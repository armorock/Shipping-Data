# Shipping-Map

Interactive Leaflet dashboard of Armorock **MHB bases shipped** (2014–2026) for the sales team. Cluster badges sum bases in the ground.

**Pin color = location precision:** red = zip (precise), orange = city (moderate), gray = state-level (approximate). **Pin shape = plant:** NV = ✕, TX = ★, FL = ●.

**Filters** (left panel): job/project/customer search, year (dropdown of checkboxes), single-state dropdown, plant, location precision. **Display toggles:** state-heat choropleth, heatmap glow, plant → destination arcs, show/hide pins.

**Presentation graphics:** KPI header tiles (total bases, states served, jobs, top state, latest-year growth), a by-year bar chart (click a bar to toggle that year), a diameter-mix donut, top-states and top-customers leaderboards (click a state row to isolate it), and an animated year timeline (play button + scrubber, cumulative 2014→2026). **Export PNG** button downloads the current map view for slide decks.

All charts/filters are computed client-side from `data.json`; the build pipeline is unchanged. Beyond Leaflet/Carto, the page loads three more CDN libraries at runtime: a US-states GeoJSON (jsDelivr) for the choropleth (degrades gracefully if unavailable), `leaflet.heat` (heatmap glow), and `dom-to-image-more` (PNG export).

Live site: [armorock.github.io/ax-fieldatlas-7q3x](https://armorock.github.io/ax-fieldatlas-7q3x/) (public repo with obscure name + noindex; org plan does not support private Pages)

## How to run

```sh
python download_geodata.py    # one-time: fetches GeoNames US.txt into data/
python build_map_data.py      # builds output/data.json, copies index.html, prints verification
python -m http.server 8741 -d output   # local preview at http://localhost:8741
```

## Inputs

- `C:\Users\JohnLeitzke\OneDrive - Armorock LLC\Documents\Desktop\NSAW All Shipping Data1.2.xlsx` — Sheet1 only (line-item shipping data)
- `..\Schooleys Shit\Jacks_Data_Improvement_Plans\output\jobcode_db.json` — fills missing city/state/zip by job code; supplies project name and contractor
- `data/US.txt` — GeoNames postal centroids (CC BY 4.0, credit shown in map attribution)

## Outputs

`output/` is the deployable site (also a standalone git repo pushed to `armorock/ax-fieldatlas-7q3x`):

- `index.html` — single-file Leaflet app (copied from the authored source at the project root)
- `data.json` — compact build: `locs` (deduped lat/lng), `jobs`, `customers`, `mhb` records `[locIdx, year, plantIdx, qty, jobIdx, prec, [[diam, count], ...]]`
- `robots.txt`

## Data rules

- Rows with Quantity = 0 or blank are excluded; MHB row count = MHB base count
- Geocoding precision: 0 = zip centroid, 1 = city centroid (fuzzy match cutoff 0.85, same state, logged at build), 2 = state centroid (blue pins, "approximate" in popup)
- Rows with no usable state are excluded and counted in the map's "Not mapped" note
- MHB verification baseline (must match build output): total 16,559 after qty=0/blank exclusion

## Redeploy after data changes

```sh
python build_map_data.py
cd output
git add data.json index.html && git commit -m "refresh data" && git push
```
