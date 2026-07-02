# Shipping-Map

Interactive Leaflet dashboard of Armorock **MHB bases shipped** (2014–present) for the sales team. Cluster badges sum bases in the ground.

**Pin color = location precision:** red = zip (precise), orange = city (moderate), gray = state-level (approximate), purple = suspect city/state (hidden by default). **Pin shape = plant:** NV = ✕, TX = ★, FL = ●.

**Filters** (left panel): job/project/customer search, year (dropdown of checkboxes), single-state dropdown, plant, location precision. **Display toggles:** state-heat choropleth, heatmap glow, plant → destination arcs, show/hide pins.

**Presentation graphics:** KPI header tiles (total bases, states served, jobs, top state, YTD growth vs prior year), a by-year bar chart (click a bar to toggle that year), a diameter-mix donut, top-states and top-customers leaderboards (click a state row to isolate it), and an animated year timeline (play button + scrubber, cumulative 2014→present). **Export PNG** button downloads the current map view for slide decks.

All charts/filters are computed client-side from `data.json`. Beyond Leaflet/Carto, the page loads three CDN libraries at runtime: US-states GeoJSON for the choropleth, `leaflet.heat` (heatmap glow), and `dom-to-image-more` (PNG export).

Live site: [armorock.github.io/ax-fieldatlas-7q3x](https://armorock.github.io/ax-fieldatlas-7q3x/) (public repo with obscure name + noindex; org plan does not support private Pages)

## How to run

```sh
python download_geodata.py    # one-time: fetches GeoNames US.txt into data/
python build_map_data.py      # builds output/data.json, copies index.html, prints verification
python -m http.server 8741 -d output   # local preview at http://localhost:8741
```

## Quarterly update

From the project root (`Shipping-Data/`):

```sh
# 1. Download fresh NetSuite exports into Combined-Database/data/:
#    - AppxShipped saved search → AppxShipped*.xls
#    - OSLocationData saved search → OSLocationData*.xls

# 2. Run the orchestrator:
python run_quarterly_update.py

# 3. Deploy:
cd Shipping-Map/output
git add data.json index.html && git commit -m "Q refresh" && git push
```

To skip ingest and just rebuild the map: `python run_quarterly_update.py --skip-ingest`

## Inputs

- `C:\Users\JohnLeitzke\OneDrive - Armorock LLC\Documents\Desktop\NSAW All Shipping Data1.2.xlsx` — Sheet1, line-item shipping history (appended quarterly by `ingest_netsuite.py`)
- `..\Schooleys Shit\Jacks_Data_Improvement_Plans\output\jobcode_db.json` — fills missing city/state/zip by job code; supplies project name and contractor
- `data/US.txt` — GeoNames postal centroids (CC BY 4.0)
- `..\Combined-Database\data\AppxShipped*.xls` and `OSLocationData*.xls` — NetSuite quarterly exports

## Outputs

`output/` is the deployable site (also a standalone git repo pushed to `armorock/ax-fieldatlas-7q3x`):

- `index.html` — single-file Leaflet app (copied from the authored source at the project root)
- `data.json` — compact build: `locs`, `jobs`, `customers`, `mhb` records
- `geo_anomalies.csv` — suspect city/state combos that couldn't be auto-corrected (for manual review)
- `robots.txt`

## Data rules

- Rows with Quantity = 0 or blank are excluded; MHB row count = MHB base count
- Geocoding precision: 0 = zip centroid, 1 = city centroid (fuzzy 0.85 cutoff, same state), 2 = state centroid, 3 = suspect city/state
- Auto-correction: if a suspect city exists in exactly one GeoNames state and no zip contradicts it, the state is corrected silently; popup shows green "State corrected" note
- Rows with no usable state are excluded and counted in the map's "Not mapped" note
- MHB verification baseline: 16,839 as of 2026-07-02 (updated each quarter in `MHB_BASELINE` dict)
