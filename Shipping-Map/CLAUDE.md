# Shipping-Map — project notes

Interactive Leaflet dashboard of Armorock MHB bases shipped (2014–present) for the sales team. Live at `armorock/ax-fieldatlas-7q3x` on GitHub Pages.

## Files

- `build_map_data.py` — reads NSAW xlsx + jobcode_db.json, geocodes, writes `output/data.json` and copies `index.html`
- `index.html` — single-file Leaflet app; all charts/filters computed client-side from `data.json`
- `download_geodata.py` — one-time fetch of GeoNames `data/US.txt`
- `output/` — deployable site, its own git repo pushed to `armorock/ax-fieldatlas-7q3x`

## data.json schema

```
meta: { generated, currentMonth, notMappedMhb, notMappedRows, badDates, years, plants }
locs: [[lat, lng], ...]
customers: [name, ...]
jobs: [jobCode, project, custIdx, contractor, city, state, zip, [[part,n],...top15], [minYr, maxYr]]
mhb: [locIdx, year, month, plantIdx, qty, jobIdx, prec, [[diam, count], ...]]
```

`month` (1–12) is included so the frontend can compute YTD comparisons (cap prior year at the same month as `currentMonth`).

## Key data rules

- Rows with Quantity = 0 or blank are excluded
- Valid plants: "Boulder City", "Sulphur Springs", "Plant City"
- Geocoding: prec 0 = zip centroid, 1 = city centroid (fuzzy cutoff 0.85, same state), 2 = state centroid
- MHB baseline total: 16,559 (build prints PASS/FAIL against `MHB_BASELINE` dict)
- `generated` date is set from `date.today()` at build time — no need to hardcode

## Growth KPI (YTD comparison)

The "YTD growth" KPI tile compares the current year's shipments (Jan–`currentMonth`) to the same window of the prior year. This avoids the mid-year dip that appears when comparing a partial year to a full prior year.

Implementation: `month` is stored on each marker's `.meta`; `updateKPI` sums `visible` markers where `year === maxYr - 1 && month <= DATA.meta.currentMonth` as the denominator.

## Redeploy

```sh
python build_map_data.py
cd output
git add data.json index.html && git commit -m "refresh data" && git push
```

## Input paths

- NSAW xlsx: `C:\Users\JohnLeitzke\OneDrive - Armorock LLC\Documents\Desktop\NSAW All Shipping Data1.2.xlsx` (Sheet1)
- jobcode_db: `..\Schooleys Shit\Jacks_Data_Improvement_Plans\output\jobcode_db.json`
- GeoNames: `data/US.txt`
