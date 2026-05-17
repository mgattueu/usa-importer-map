import openpyxl, re, json, math, random
import pgeocode
from collections import Counter

wb = openpyxl.load_workbook(r'C:\Users\mgatt\Downloads\ImportYeti-PowerQuery-05-16-2026.xlsx', read_only=True)
ws = wb['Data']
rows = list(ws.iter_rows(values_only=True))
print(f"Rows: {len(rows)-1}")

nomi_us = pgeocode.Nominatim('us')
nomi_ca = pgeocode.Nominatim('ca')
_zip_cache = {}

def zip_coords(z, country='us'):
    key = (z, country)
    if key not in _zip_cache:
        try:
            r = (nomi_us if country == 'us' else nomi_ca).query_postal_code(z)
            lat, lng = float(r['latitude']), float(r['longitude'])
            _zip_cache[key] = None if (math.isnan(lat) or math.isnan(lng)) else (lat, lng)
        except:
            _zip_cache[key] = None
    return _zip_cache[key]

STATE_LL = {
    'AL':(32.36,-86.30),'AK':(64.20,-153.37),'AZ':(34.04,-111.09),'AR':(34.80,-92.20),
    'CA':(36.78,-119.42),'CO':(39.06,-105.31),'CT':(41.60,-72.75),'DE':(38.91,-75.53),
    'FL':(27.66,-81.52),'GA':(32.17,-82.90),'HI':(19.90,-155.58),'ID':(44.24,-114.48),
    'IL':(40.35,-88.99),'IN':(39.85,-86.26),'IA':(42.01,-93.21),'KS':(38.53,-96.73),
    'KY':(37.84,-84.27),'LA':(30.98,-91.96),'ME':(45.25,-69.45),'MD':(39.05,-76.64),
    'MA':(42.24,-71.53),'MI':(43.33,-84.54),'MN':(45.69,-93.90),'MS':(32.75,-89.67),
    'MO':(38.46,-92.29),'MT':(46.88,-110.36),'NE':(41.49,-99.90),'NV':(38.31,-117.06),
    'NH':(43.45,-71.56),'NJ':(40.06,-74.41),'NM':(34.84,-106.25),'NY':(42.17,-74.95),
    'NC':(35.63,-79.81),'ND':(47.55,-101.00),'OH':(40.39,-82.76),'OK':(35.57,-96.93),
    'OR':(44.57,-122.07),'PA':(40.59,-77.21),'RI':(41.68,-71.51),'SC':(33.84,-81.16),
    'SD':(43.97,-99.90),'TN':(35.75,-86.69),'TX':(31.05,-97.56),'UT':(39.32,-111.09),
    'VT':(44.05,-72.71),'VA':(37.43,-78.66),'WA':(47.40,-121.49),'WV':(38.49,-80.95),
    'WI':(43.78,-88.79),'WY':(42.76,-107.30),'DC':(38.91,-77.02),
}
US_STATES = set(STATE_LL.keys())

def extract_zip(text):
    s = str(text).upper()
    m = re.search(r'\b([A-Z]\d[A-Z]\s*\d[A-Z]\d)\b', s)
    if m: return m.group(1).replace(' ',''), 'ca'
    m = re.search(r'(\d{5})(?:\d{4})?', s)
    if m: return m.group(1), 'us'
    return None, 'us'

def extract_state(text):
    s = str(text).upper()
    found = re.findall(r'\b([A-Z]{2})\b', s)
    for code in reversed(found):
        if code in US_STATES: return code
    return ''

def split_addresses(text):
    """Split a multi-address cell into individual address strings."""
    if not text: return []
    s = str(text).strip()
    # Split on boundaries between one address ending and the next starting.
    # Addresses typically end with: zip, "Us", "USA", "United States", "Canada"
    # then the next starts with a digit or capital word.
    parts = re.split(
        r'(?:(?<=\d{5})|(?<=Us)|(?<=USA)|(?<=Canada)|(?<=States))\s*,\s*(?=[A-Z0-9])',
        s, flags=re.IGNORECASE
    )
    return [p.strip() for p in parts if len(p.strip()) > 10]

def geocode_address(text):
    """Return (lat, lng, method) for a single address string."""
    z, country = extract_zip(text)
    if z:
        c = zip_coords(z, country)
        if c: return c[0], c[1], 'zip'
    st = extract_state(text)
    if st in STATE_LL:
        return STATE_LL[st][0], STATE_LL[st][1], 'state'
    return None, None, None

def clean_addr(text):
    """Human-readable version of an address."""
    s = str(text).strip()
    s = re.sub(r'\s+', ' ', s)
    return s[:120]

# ── Categorise ──────────────────────────────────────────────────────
LOGISTICS_KW = [
    'logistics','freight','maritime','customs broker','forwarding','airlift',
    'interglobo','oec ','oec freight','pegasus maritime','firstlift','patagon',
    'comfy logistics','brightway','ctc logistics','relay logistics',
    'worldwide logistics','imperative logistics','apex maritime','cil freight',
    'umax shipping','interfreight','leela logistics','mts logistics',
    'blackstone shipping','ups ocean','penguin shipping',
    'worldwide logistics partners','american global freight','oocl',
    'c&l container','rtw logistics','jenson logistics','topocean',
    'city ocean','transmarine cargo','em lines ltd','famous pacific',
    'pan pacific express','ctl lax','car go worldwide','amass global',
    'affinity shipping','cts global supply','dsv air','cms shipping',
    'western overseas','ascend express','sea dominion','binex line',
    'orient express container','crossea shipping','seahorse container',
    'seax trade','seko worldwide','hecny','winfar intl','stable enterprise',
    'permeco','best global management','tanera transport','jdy international',
    'db group america','jupiter international usa','general noli',
    'ftl plus','world class shipping','skytrans','jr global',
    'del corona','gh trans','speedier logistic','oriental air transport',
    'advantage group intl','shipco transport','comage container',
    'glenrock international','a j worldwide','1up cargo','gran trade',
    'express consolidation','aetos cargo','vg enterprises','tlss inc',
    'aprile usa','troy container','city ocean intl','magellan shipping',
    'gateway international llc','norman krieger','john s connor','unitrans',
    'ch robinson','alliance trade','tql global','midwest transatlantic',
    'columbus customhouse','sjlt usa','triumph express','traffic tech',
    'harvest logistic','locher evers','seair global','rohlig',
    'rainbow import','fcg global','carnevale','acrocargo',
    'master logistix','bbe expediting','ecu worldwide',
    'cole international','canaan transport','booking union','ultra air cargo',
    'savino del bene','sar logisolutions','db shipping',
    'vnft international','associated import corp','dahnay logistics',
    'worldwide logistic partners','worldwide logistics ltd',
    'oec freight','oec miami','interglobo north america',
]

def categorize(name, vol):
    """Everyone importing is a competitor. < 10 TEU = potential customer."""
    n = name.lower()
    if 'kani international' in n or 'slab planet' in n: return 'SELF'
    for k in LOGISTICS_KW:
        if k in n: return 'LOGISTICS'
    # volume-based split: small importers are potential customers
    if vol < 10:
        return 'CUSTOMER'
    return 'COMPETITOR'

# ── Process rows ────────────────────────────────────────────────────
markers   = []
geo_stats = Counter()

random.seed(42)

def jitter(lat, lng, used, radius=0.18):
    for _ in range(20):
        angle = random.uniform(0, 2*math.pi)
        r = random.uniform(0.05, radius)
        p = (round(lat + r*math.sin(angle), 5), round(lng + r*math.cos(angle), 5))
        if p not in used:
            used.add(p)
            return p
    used.add((lat, lng))
    return lat, lng

used_coords = set()

for row in rows[1:]:
    vol    = row[0] or 0
    name   = str(row[2] or '').strip()
    notify = str(row[4] or '').strip()
    addr2  = str(row[5] or '').strip()
    if not name: continue

    cat = categorize(name, vol)

    # Collect all address strings from both columns
    all_raw = split_addresses(notify) + split_addresses(addr2)

    # Deduplicate by zip code — keep first occurrence of each zip
    seen_zips = set()
    unique_addrs = []
    for raw in all_raw:
        z, country = extract_zip(raw)
        key = (z or raw[:30], country)
        if key not in seen_zips:
            seen_zips.add(key)
            lat, lng, method = geocode_address(raw)
            if lat is not None:
                unique_addrs.append({
                    'text': clean_addr(raw),
                    'lat': lat,
                    'lng': lng,
                    'method': method,
                })

    if not unique_addrs:
        geo_stats['no_coords'] += 1
        continue

    # Primary address = first one with a zip (most reliable)
    primary = unique_addrs[0]
    secondary = unique_addrs[1:]  # genuinely different locations

    geo_stats[primary['method']] += 1

    plat, plng = jitter(primary['lat'], primary['lng'], used_coords)

    markers.append({
        'name': name,
        'cat' : cat,
        'vol' : int(vol),
        'lat' : plat,
        'lng' : plng,
        'geo' : primary['method'],
        'primary_addr': primary['text'],
        'secondary': secondary,   # list of {text, lat, lng, method}
    })

print(f"\nGeocoding: {dict(geo_stats)}")
print(f"Mapped: {len(markers)}  |  No coords: {geo_stats['no_coords']}")
print(f"Companies with 2+ locations: {sum(1 for m in markers if m['secondary'])}")
print(f"Categories: {Counter(m['cat'] for m in markers)}")

# ── HTML ─────────────────────────────────────────────────────────────
mj = json.dumps(markers, ensure_ascii=False)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Kani – Granite Monument Import Map</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css"/>
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#e2e8f0;height:100vh;display:flex;flex-direction:column}}
#hdr{{padding:10px 16px;background:#1e293b;display:flex;align-items:center;gap:12px;border-bottom:2px solid #334155;flex-shrink:0;flex-wrap:wrap}}
#hdr b{{font-size:1rem;color:#f1f5f9}}
#hdr small{{font-size:.71rem;color:#64748b}}
.leg{{display:flex;gap:12px;margin-left:auto;flex-wrap:wrap;align-items:center}}
.li{{display:flex;align-items:center;gap:5px;font-size:.74rem;color:#cbd5e1}}
.dot{{width:10px;height:10px;border-radius:50%;flex-shrink:0}}
.d-comp{{background:#ef4444}}.d-buy{{background:#22c55e}}
.d-inv{{background:#f59e0b}}.d-log{{background:#94a3b8}}.d-self{{background:#3b82f6}}

#bar{{padding:6px 14px;background:#1e293b;display:flex;gap:7px;flex-wrap:wrap;align-items:center;border-bottom:1px solid #334155;flex-shrink:0}}
.btn{{padding:3px 11px;border-radius:13px;border:none;cursor:pointer;font-size:.74rem;font-weight:600;transition:opacity .15s}}
.btn.off{{opacity:.28}}
.b-all{{background:#475569;color:#fff}}.b-comp{{background:#dc2626;color:#fff}}
.b-cust{{background:#22c55e;color:#fff}}.b-log{{background:#6b7280;color:#fff}}
.b-self{{background:#3b82f6;color:#fff}}.b-lbl{{background:#0f172a;color:#94a3b8;border:1px solid #334155}}
#search{{margin-left:auto;padding:4px 12px;border-radius:13px;border:1px solid #475569;background:#0f172a;color:#e2e8f0;font-size:.74rem;width:175px}}

/* Focus mode banner */
#focus-bar{{display:none;padding:7px 16px;background:#7c3aed;color:#fff;font-size:.82rem;font-weight:600;align-items:center;gap:12px;flex-shrink:0}}
#focus-bar button{{padding:3px 12px;border-radius:12px;border:none;background:rgba(255,255,255,.25);color:#fff;cursor:pointer;font-weight:700;font-size:.76rem}}
#focus-bar button:hover{{background:rgba(255,255,255,.4)}}

#map{{flex:1;position:relative}}

/* Focus location markers */
.loc-pin{{
  display:flex;align-items:center;justify-content:center;
  width:32px;height:32px;border-radius:50%;
  font-size:.72rem;font-weight:700;color:#fff;
  border:2.5px solid rgba(255,255,255,.7);
  box-shadow:0 2px 10px rgba(0,0,0,.6);
}}
.loc-primary{{background:#3b82f6}}
.loc-secondary{{background:#7c3aed}}

/* Side panel */
#panel{{
  position:absolute;top:10px;left:10px;z-index:1000;
  background:#1e293b;border:1px solid #334155;border-radius:10px;
  width:290px;max-height:calc(100vh - 130px);
  overflow-y:auto;box-shadow:0 6px 28px rgba(0,0,0,.7);
  display:none;
}}
#panel-inner{{padding:14px}}
.pn{{font-weight:700;font-size:.95rem;margin-bottom:5px;line-height:1.3}}
.pb{{display:inline-block;padding:2px 9px;border-radius:10px;font-size:.69rem;font-weight:700;margin-bottom:7px}}
.bc{{background:#7f1d1d;color:#fca5a5}}.bb{{background:#14532d;color:#86efac}}
.bl{{background:#334155;color:#cbd5e1}}.bs{{background:#1e3a8a;color:#93c5fd}}
.pv{{font-size:.8rem;font-weight:600;color:#e2e8f0;margin-bottom:10px}}
.addr-label{{font-size:.64rem;text-transform:uppercase;letter-spacing:.5px;color:#64748b;margin-bottom:5px}}
.addr-item{{display:flex;align-items:flex-start;gap:8px;padding:6px 8px;border-radius:6px;margin-bottom:5px;cursor:pointer;transition:background .15s;border:1px solid transparent}}
.addr-item:hover,.addr-item.active{{background:rgba(255,255,255,.07);border-color:#334155}}
.addr-num{{width:22px;height:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:.66rem;font-weight:700;flex-shrink:0;margin-top:1px}}
.addr-primary-num{{background:#3b82f6;color:#fff}}
.addr-secondary-num{{background:#7c3aed;color:#fff}}
.addr-text{{font-size:.76rem;color:#cbd5e1;line-height:1.38}}
.addr-geo{{font-size:.63rem;color:#475569;margin-top:2px}}
.addr-warn{{color:#f59e0b}}

/* Permanent label on every marker */
.plabel{{background:rgba(15,23,42,.82)!important;border:1px solid #334155!important;color:#e2e8f0!important;font-size:.7rem!important;padding:2px 6px!important;border-radius:4px!important;box-shadow:0 1px 4px rgba(0,0,0,.5)!important;white-space:nowrap!important;pointer-events:none!important}}
.plabel::before{{display:none!important}}
.mlabel{{display:flex;gap:4px;align-items:center}}
.mteu{{color:#94a3b8;font-size:.65rem}}
/* Focus address labels */
.ftip{{background:#1e293b!important;border:1px solid #7c3aed!important;color:#e2e8f0!important;font-size:.73rem!important;padding:3px 8px!important;border-radius:4px!important;box-shadow:0 2px 6px rgba(0,0,0,.5)!important;white-space:normal!important;max-width:220px}}
.ftip::before{{border-right-color:#7c3aed!important}}

#stats{{position:absolute;bottom:20px;right:10px;z-index:999;background:#1e293b;border:1px solid #334155;border-radius:8px;padding:8px 12px;font-size:.72rem;min-width:150px}}
#stats h4{{color:#64748b;text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;font-size:.67rem}}
.sr{{display:flex;justify-content:space-between;gap:10px;margin:2px 0}}
.sl{{color:#94a3b8}}.sv{{font-weight:700;color:#f1f5f9}}
</style>
</head>
<body>
<div id="hdr">
  <div><b>Kani – Granite Monument Import Map</b><br><small>ImportYeti May-16-2026 · {len(markers)} companies · click any pin to focus</small></div>
  <div class="leg">
    <div class="li"><div class="dot" style="background:#dc2626"></div>Competitor (≥10 TEU)</div>
    <div class="li"><div class="dot" style="background:#22c55e"></div>Potential Customer (&lt;10 TEU)</div>
    <div class="li"><div class="dot" style="background:#6b7280"></div>🚚 Logistics</div>
    <div class="li"><div class="dot" style="background:#3b82f6"></div>Kani</div>
    <div class="li" style="font-size:.68rem;color:#64748b">darker red = higher volume</div>
  </div>
</div>
<div id="bar">
  <button class="btn b-all"  onclick="showAll()">All</button>
  <button class="btn b-comp" id="btn-COMPETITOR" onclick="tog('COMPETITOR')">Competitors</button>
  <button class="btn b-cust" id="btn-CUSTOMER"   onclick="tog('CUSTOMER')">Potential Customers</button>
  <button class="btn b-log"  id="btn-LOGISTICS"  onclick="tog('LOGISTICS')">Logistics</button>
  <button class="btn b-self" id="btn-SELF"        onclick="tog('SELF')">Kani</button>
  <button class="btn b-lbl" id="btn-labels" onclick="toggleLabels()">Labels ON</button>
  <input id="search" type="text" placeholder="Search company…" oninput="doSearch(this.value)">
</div>
<div id="focus-bar">
  <span id="focus-name"></span>
  <button onclick="exitFocus()">← Back to all companies</button>
</div>
<div id="map">
  <div id="panel">
    <div id="panel-inner"></div>
  </div>
</div>
<div id="stats">
  <h4>Visible</h4>
  <div class="sr"><span class="sl">Competitors</span>   <span class="sv" id="s-c">-</span></div>
  <div class="sr"><span class="sl">Pot. Customers</span><span class="sv" id="s-k">-</span></div>
  <div class="sr"><span class="sl">Logistics</span>     <span class="sv" id="s-l">-</span></div>
  <div class="sr"><span class="sl">Total pins</span>    <span class="sv" id="s-t">-</span></div>
  <div style="margin-top:6px;font-size:.64rem;color:#7c3aed">🟣 violet = state-level pin</div>
</div>

<script>
const DATA={mj};
// Color by volume for competitors (red=large, orange=medium)
function compColor(vol){{
  if(vol>=100) return '#b91c1c';   // dark red
  if(vol>=50)  return '#dc2626';   // red
  if(vol>=20)  return '#ea580c';   // red-orange
  return '#f97316';                // orange
}}
const COLORS={{COMPETITOR:'#dc2626',CUSTOMER:'#22c55e',LOGISTICS:'#6b7280',SELF:'#3b82f6'}};
const LABELS={{COMPETITOR:'Competitor',CUSTOMER:'Potential Customer (<10 TEU)',LOGISTICS:'Logistics / Freight',SELF:'Kani / Slab Planet'}};
const BADGE ={{COMPETITOR:'bc',CUSTOMER:'bb',LOGISTICS:'bl',SELF:'bs'}};

const map=L.map('map',{{center:[38,-96],zoom:4}});
L.tileLayer('https://{{s}}.basemaps.cartocdn.com/dark_all/{{z}}/{{x}}/{{y}}{{r}}.png',{{attribution:'&copy; OpenStreetMap &copy; CartoDB',maxZoom:19}}).addTo(map);

// ── Icons ──────────────────────────────────────────────────────────
function makeClusterIcon(cat,vol,geo){{
  if(cat==='LOGISTICS')
    return L.divIcon({{html:'<div style="font-size:15px;line-height:1;filter:grayscale(1) opacity(.7)">🚚</div>',className:'',iconSize:[20,18],iconAnchor:[10,9]}});
  // State-level geocoded pins are violet
  if(geo==='state') return makeStateIcon(vol);
  const c=cat==='COMPETITOR'?compColor(vol):(COLORS[cat]||'#888');
  const r=cat==='SELF'?13:Math.max(6,Math.min(6+Math.sqrt(vol)*1.8,22));
  const s=r*2+4;
  const fs=Math.max(7,Math.min(10,r));
  const lbl=vol>0?`<text x="${{s/2}}" y="${{s/2+fs*.38}}" text-anchor="middle" font-size="${{fs}}" font-weight="700" fill="white" font-family="Segoe UI,sans-serif">${{vol}}</text>`:'';
  return L.divIcon({{html:`<svg xmlns="http://www.w3.org/2000/svg" width="${{s}}" height="${{s}}"><circle cx="${{s/2}}" cy="${{s/2}}" r="${{r}}" fill="${{c}}" fill-opacity=".9" stroke="white" stroke-width="1.5"/>${{lbl}}</svg>`,className:'',iconSize:[s,s],iconAnchor:[s/2,s/2]}});
}}

function makeStateIcon(vol){{
  const r=Math.max(6,Math.min(6+Math.sqrt(vol)*1.8,22));
  const s=r*2+4; const fs=Math.max(7,Math.min(10,r));
  const lbl=vol>0?`<text x="${{s/2}}" y="${{s/2+fs*.38}}" text-anchor="middle" font-size="${{fs}}" font-weight="700" fill="white" font-family="Segoe UI,sans-serif">${{vol}}</text>`:'';
  return L.divIcon({{html:`<svg xmlns="http://www.w3.org/2000/svg" width="${{s}}" height="${{s}}"><circle cx="${{s/2}}" cy="${{s/2}}" r="${{r}}" fill="#7c3aed" fill-opacity=".85" stroke="white" stroke-width="1.5" stroke-dasharray="4,2"/>${{lbl}}</svg>`,className:'',iconSize:[s,s],iconAnchor:[s/2,s/2]}});
}}

function makeSecondaryIcon(cat,vol){{
  const c=cat==='COMPETITOR'?compColor(vol):(COLORS[cat]||'#888');
  const r=5; const s=r*2+4;
  return L.divIcon({{html:`<svg xmlns="http://www.w3.org/2000/svg" width="${{s}}" height="${{s}}"><circle cx="${{s/2}}" cy="${{s/2}}" r="${{r}}" fill="${{c}}" fill-opacity=".55" stroke="white" stroke-width="1.5" stroke-dasharray="3,2"/></svg>`,className:'',iconSize:[s,s],iconAnchor:[s/2,s/2]}});
}}

function makeFocusIcon(idx){{
  const cls=idx===1?'loc-primary':'loc-secondary';
  return L.divIcon({{html:`<div class="loc-pin ${{cls}}">${{idx}}</div>`,className:'',iconSize:[32,32],iconAnchor:[16,16]}});
}}

// ── Cluster groups (normal mode) ───────────────────────────────────
const CATS=['COMPETITOR','CUSTOMER','LOGISTICS','SELF'];
const clusters={{}},active={{}};
const allItems=[];   // one entry per map pin (may be multiple per company)

CATS.forEach(cat=>{{
  active[cat]=true;
  clusters[cat]=L.markerClusterGroup({{
    maxClusterRadius:45,
    iconCreateFunction:function(cl){{
      const n=cl.getChildCount(),c=COLORS[cat];
      return L.divIcon({{html:`<div style="background:${{c}};opacity:.82;width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:11px;border:2px solid rgba(255,255,255,.3)">${{n}}</div>`,className:'',iconSize:[32,32],iconAnchor:[16,16]}});
    }}
  }});
  map.addLayer(clusters[cat]);
}});

// Create one marker per location (primary + all secondaries) for every company
DATA.forEach(d=>{{
  const locs=[{{text:d.primary_addr,lat:d.lat,lng:d.lng,method:d.geo,isPrimary:true}},
              ...(d.secondary||[]).map(s=>{{return {{...s,isPrimary:false}};}})];;

  locs.forEach((loc,i)=>{{
    const isPrimary=loc.isPrimary;
    const icon=isPrimary
      ? makeClusterIcon(d.cat, d.vol, loc.method)
      : makeSecondaryIcon(d.cat, d.vol);
    const m=L.marker([loc.lat,loc.lng],{{icon}});
    // Permanent label: short name + TEU (skip logistics)
    if(d.cat!=='LOGISTICS'){{
      const shortName=d.name.length>28?d.name.substring(0,26)+'…':d.name;
      const labelHtml=`<span class="mlabel">${{shortName}}<span class="mteu"> ${{d.vol}}t</span></span>`;
      m.bindTooltip(labelHtml,{{permanent:true,direction:'right',offset:[8,0],className:'plabel'}});
    }}
    m.on('click',()=>enterFocus(d));
    m._d=d; m._vis=true;
    clusters[d.cat].addLayer(m);
    allItems.push(m);
  }});
}});

// ── Focus mode ─────────────────────────────────────────────────────
let focusLayers=[];
let inFocus=false;
let _focusBounds=[];

function enterFocus(d){{
  inFocus=true;
  // Hide all cluster groups
  CATS.forEach(cat=>map.removeLayer(clusters[cat]));
  // Clear old focus markers
  focusLayers.forEach(l=>map.removeLayer(l));
  focusLayers=[];

  // Build ordered list: primary first, then secondaries
  const locs=[{{text:d.primary_addr,lat:d.lat,lng:d.lng,method:d.geo}},...(d.secondary||[])];

  // Place numbered pins for every location, with permanent address labels
  const bounds=[];
  locs.forEach((loc,i)=>{{
    const short=loc.text.replace(/[ ]+(Us|USA|United States|Canada).*$/i,'').trim();
    const m=L.marker([loc.lat,loc.lng],{{icon:makeFocusIcon(i+1),zIndexOffset:2000}});
    m.bindTooltip(`<b>#${{i+1}}</b> ${{short}}`,{{permanent:true,direction:'right',offset:[14,0],className:'ftip'}});
    m.addTo(map);
    focusLayers.push(m);
    bounds.push([loc.lat,loc.lng]);
  }});
  _focusBounds=bounds;

  // Fit map to show ALL locations
  if(bounds.length===1) map.setView(bounds[0],13,{{animate:true}});
  else map.fitBounds(bounds,{{padding:[80,80],animate:true}});

  // Show focus bar
  document.getElementById('focus-bar').style.display='flex';
  document.getElementById('focus-name').textContent=d.name;

  // Build side panel
  const BADGE_C=BADGE[d.cat];
  let html=`<div class="pn">${{d.name}}</div>
    <span class="pb ${{BADGE_C}}">${{LABELS[d.cat]}}</span>
    <div class="pv">Volume: ${{d.vol}} TEU</div>
    <div class="addr-label">${{locs.length}} location${{locs.length>1?'s':''}}</div>`;

  locs.forEach((loc,i)=>{{
    const numCls=i===0?'addr-primary-num':'addr-secondary-num';
    const geoTag=loc.method==='state'?'<span class="addr-warn">⚠ state-level pin</span>':'✓ zip-geocoded';
    html+=`<div class="addr-item" id="ai${{i}}" onclick="panToLoc(${{loc.lat}},${{loc.lng}},${{i}})">
      <div class="addr-num ${{numCls}}">${{i+1}}</div>
      <div>
        <div class="addr-text">${{loc.text}}</div>
        <div class="addr-geo">${{geoTag}}</div>
      </div>
    </div>`;
  }});

  document.getElementById('panel-inner').innerHTML=html;
  document.getElementById('panel').style.display='block';

  updateStats();
}}

function panToLoc(lat,lng,idx){{
  // Highlight selected row
  document.querySelectorAll('.addr-item').forEach((el,i)=>el.classList.toggle('active',i===idx));
  // Zoom to show this location AND all others together
  if(_focusBounds.length<=1){{
    map.setView([lat,lng],14,{{animate:true}});
  }} else {{
    map.fitBounds(_focusBounds,{{padding:[80,80],animate:true}});
  }}
}}

function exitFocus(){{
  inFocus=false;
  focusLayers.forEach(l=>map.removeLayer(l));
  focusLayers=[];
  document.getElementById('focus-bar').style.display='none';
  document.getElementById('panel').style.display='none';
  // Restore clusters
  CATS.forEach(cat=>{{if(active[cat]) map.addLayer(clusters[cat]);}});
  map.setView([38,-96],4,{{animate:true}});
  updateStats();
}}

// ── Stats ──────────────────────────────────────────────────────────
function updateStats(){{
  if(inFocus){{
    ['s-c','s-k','s-l','s-t'].forEach(id=>document.getElementById(id).textContent='—');
    document.getElementById('s-t').textContent='focus';
    return;
  }}
  let comp=0,cust=0,log=0,tot=0;
  allItems.forEach(m=>{{if(!m._vis)return;tot++;const c=m._d.cat;if(c==='COMPETITOR')comp++;else if(c==='CUSTOMER')cust++;else if(c==='LOGISTICS')log++;}});
  document.getElementById('s-c').textContent=comp;
  document.getElementById('s-k').textContent=cust;
  document.getElementById('s-l').textContent=log;
  document.getElementById('s-t').textContent=tot;
}}
allItems.forEach(m=>m._vis=true); updateStats();

// ── Filters (only work in normal mode) ────────────────────────────
function tog(cat){{
  if(inFocus) return;
  active[cat]=!active[cat];
  document.getElementById('btn-'+cat).classList.toggle('off',!active[cat]);
  if(active[cat]) map.addLayer(clusters[cat]); else map.removeLayer(clusters[cat]);
  allItems.forEach(m=>{{if(m._d.cat===cat)m._vis=active[cat];}});
  updateStats();
}}
function showAll(){{
  if(inFocus){{exitFocus();return;}}
  CATS.forEach(cat=>{{active[cat]=true;const b=document.getElementById('btn-'+cat);if(b)b.classList.remove('off');map.addLayer(clusters[cat]);}});
  allItems.forEach(m=>m._vis=true); updateStats();
}}
function doSearch(q){{
  if(inFocus) exitFocus();
  q=q.toLowerCase().trim();
  CATS.forEach(cat=>clusters[cat].clearLayers());
  allItems.forEach(m=>{{
    const d=m._d;
    const match=!q||d.name.toLowerCase().includes(q)||d.state.toLowerCase()===q;
    m._vis=match&&active[d.cat];
    if(m._vis) clusters[d.cat].addLayer(m);
  }});
  updateStats();
}}

// ── Label toggle ───────────────────────────────────────────────────
let labelsOn=true;
function toggleLabels(){{
  labelsOn=!labelsOn;
  document.getElementById('btn-labels').textContent=labelsOn?'Labels ON':'Labels OFF';
  document.getElementById('btn-labels').style.color=labelsOn?'#e2e8f0':'#94a3b8';
  // Show/hide all permanent tooltips via CSS
  const style=document.getElementById('label-style')||document.createElement('style');
  style.id='label-style';
  style.textContent=labelsOn?'':'.plabel{{display:none!important}}';
  document.head.appendChild(style);
}}

// Pressing Escape also exits focus
document.addEventListener('keydown',e=>{{if(e.key==='Escape'&&inFocus)exitFocus();}});
</script>
</body>
</html>"""

out = r'C:\Users\mgatt\Downloads\Kani_Map_May2026.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"\nSaved: {out}")
