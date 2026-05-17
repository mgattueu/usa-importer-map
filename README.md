# Kani Market Map

Interactive map of granite monument importers in the US, built from ImportYeti data.

## Files

| File | Purpose |
|---|---|
| `build_map_v4.py` | Main script — geocodes ImportYeti xlsx and generates the HTML map |
| `build_visit_plan.py` | Generates the Excel visit plan with color-coded tabs |
| `analyze_importyeti.py` | Regional market analysis (text output) |
| `Kani_Map_May2026.html` | The map — open in any browser, works offline |
| `Kani_VisitPlan_May2026.xlsx` | Excel visit plan |

## How to rebuild the map

1. Download latest ImportYeti export to `Downloads/`
2. Edit the filename at the top of `build_map_v4.py`
3. Run:
   ```
   python build_map_v4.py
   ```

## Map legend

- 🔴 **Red/Orange circles** — Competitors (≥10 TEU imported). Darker = higher volume.
- 🟢 **Green circles** — Potential customers (<10 TEU)
- 🟣 **Violet dashed** — State-level pin only (no zip code in ImportYeti data)
- 🚚 **Grey truck** — Logistics / freight forwarders
- 🔵 **Blue** — Kani / Slab Planet

## Map interactions

- **Click any pin** → focus mode: hides all others, shows all locations for that company
- **← Back** or **Escape** → return to full view
- **Labels ON/OFF** toggle in toolbar
- Filter buttons to show/hide each category
- Search box to find a specific company
