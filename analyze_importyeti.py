import openpyxl, re
from collections import defaultdict

wb = openpyxl.load_workbook(r'C:\Users\mgatt\Downloads\ImportYeti-PowerQuery-05-16-2026.xlsx', read_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))

def extract_state(addr):
    if not addr: return 'UNKNOWN'
    m = re.findall(r'\b([A-Z]{2})\s*\d{5}\b', str(addr).upper())
    return m[0] if m else 'UNKNOWN'

def extract_city(addr):
    if not addr: return ''
    addr = str(addr)
    first = addr.split(',')[0].strip()
    m = re.search(r'([A-Za-z\s]+)\s+([A-Z]{2})\s*\d{5}', first)
    if m:
        return m.group(1).strip()
    return ''

def extract_full_addr(addr):
    if not addr: return ''
    return str(addr).split(',')[0].strip()[:80]

LOGISTICS_KEYWORDS = [
    'logistics','freight','maritime','customs broker','forwarding','airlift',
    'interglobo','oec ','oec freight','pegasus maritime','firstlift','patagon',
    'comfy','brightway','ctc logistics','relay logistics','worldwide logistics',
    'imperative logistics','apex maritime','cil freight','umax shipping',
    'interfreight','leela logistics','mts logistics','blackstone shipping',
    'ups ocean','penguin shipping','worldwide logistics partners',
    'american global freight','oocl','c&l container','rtw logistics',
    'jenson logistics'
]

def categorize(name, product):
    n = name.lower()
    p = (product or '').lower()
    if 'kani international' in n: return 'SELF'
    for k in LOGISTICS_KEYWORDS:
        if k in n: return 'LOGISTICS'
    if any(x in n for x in ['quarr']):
        return 'COMPETITOR'
    if any(x in n for x in ['granite','stone','monument','marble','memorial','nustone','cosentino']):
        return 'COMPETITOR'
    if any(x in n for x in ['york group','blast craft','harris enterprises','buchanan group',
                              'boc international','kerry apex','sabbow','alpi']):
        return 'COMPETITOR'
    return 'POTENTIAL_CUSTOMER'

data = []
for r in rows[1:]:
    vol = r[0] or 0
    teu = r[1] or 0
    name = r[2] or ''
    product = r[3] or ''
    notify_addr = r[4] or ''
    state = extract_state(notify_addr)
    city = extract_city(notify_addr)
    cat = categorize(name, product)
    addr = extract_full_addr(notify_addr)
    data.append({'vol': vol, 'teu': teu, 'name': name, 'state': state, 'city': city, 'cat': cat, 'addr': addr})

REGIONS = {
    'Northeast': ['NJ','NY','MA','CT','RI','VT','NH','ME','PA','MD','DE'],
    'Southeast': ['GA','NC','SC','FL','VA','AL','MS','TN','KY','LA','AR'],
    'Midwest':   ['IL','OH','IN','MI','WI','MN','IA','MO','KS'],
    'South-TX':  ['TX','OK'],
    'West':      ['CA','WA','OR','NV','UT','CO','AZ','ID','MT','WY'],
}
state_to_region = {}
for reg, states in REGIONS.items():
    for s in states:
        state_to_region[s] = reg

by_region = defaultdict(lambda: {'comp_vol':0,'comp_cnt':0,'cust_vol':0,'cust_cnt':0,'total_vol':0})

for d in data:
    if d['cat'] in ('SELF','LOGISTICS'): continue
    reg = state_to_region.get(d['state'], 'Other')
    by_region[reg]['total_vol'] += d['vol']
    if d['cat'] == 'COMPETITOR':
        by_region[reg]['comp_vol'] += d['vol']
        by_region[reg]['comp_cnt'] += 1
    else:
        by_region[reg]['cust_vol'] += d['vol']
        by_region[reg]['cust_cnt'] += 1

print("=" * 70)
print("REGIONAL MARKET ANALYSIS (excl. logistics & Kani itself)")
print("=" * 70)
print(f"{'REGION':<12} {'TOTAL VOL':>9} {'COMP VOL':>9} {'COMP #':>7} {'CUST VOL':>9} {'CUST #':>7}")
print("-" * 70)
order = ['Northeast','Southeast','Midwest','South-TX','West','Other']
for reg in order:
    r = by_region[reg]
    print(f"{reg:<12} {r['total_vol']:>9} {r['comp_vol']:>9} {r['comp_cnt']:>7} {r['cust_vol']:>9} {r['cust_cnt']:>7}")

print()
print("=" * 70)
print("COMPETITORS - SKIP VISITING (sorted by volume)")
print("=" * 70)
comps = [d for d in data if d['cat'] == 'COMPETITOR']
comps.sort(key=lambda x: -x['vol'])
for d in comps:
    city_state = f"{d['city']}, {d['state']}" if d['city'] else d['state']
    print(f"  vol={d['vol']:4}  {d['name']:<45} {city_state}")

print()
print("=" * 70)
print("POTENTIAL CUSTOMERS - VISIT (sorted by state then volume)")
print("=" * 70)
custs = [d for d in data if d['cat'] == 'POTENTIAL_CUSTOMER']
custs.sort(key=lambda x: (x['state'], -x['vol']))
cur_state = None
for d in custs:
    if d['state'] != cur_state:
        print(f"\n  --- {d['state']} ---")
        cur_state = d['state']
    city_state = f"{d['city']}, {d['state']}" if d['city'] else d['state']
    print(f"  vol={d['vol']:4}  {d['name']:<45} {city_state}")

print()
print("=" * 70)
print("WAREHOUSE EXPANSION OPPORTUNITY SUMMARY")
print("=" * 70)
for reg in order:
    r = by_region[reg]
    opp = r['cust_vol']
    comp = r['comp_vol']
    ratio = (opp / comp * 100) if comp > 0 else 999
    flag = '<== HIGH OPPORTUNITY' if opp > 50 and ratio > 30 else ''
    print(f"  {reg:<12}  customer vol={opp:>4}  competitor vol={comp:>4}  ratio={ratio:.0f}%  {flag}")
