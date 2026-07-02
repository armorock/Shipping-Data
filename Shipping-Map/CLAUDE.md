# Shipping-Map — project notes

Interactive Leaflet dashboard of Armorock MHB bases shipped (2014–present) for the sales team. Live at `armorock/ax-fieldatlas-7q3x` on GitHub Pages.

## Files

- `build_map_data.py` — reads NSAW xlsx + jobcode_db.json, geocodes, writes `output/data.json` and copies `index.html`
- `ingest_netsuite.py` — quarterly NetSuite ingest: joins AppxShipped + OSLocation XLS on `Shipment Item Transaction`, auto-detects NSAW cutoff date, appends new rows to NSAW xlsx, writes `output/geo_anomalies.csv`
- `index.html` — single-file Leaflet app; all charts/filters computed client-side from `data.json`
- `download_geodata.py` — one-time fetch of GeoNames `data/US.txt`
- `output/` — deployable site, its own git repo pushed to `armorock/ax-fieldatlas-7q3x`
- `../run_quarterly_update.py` — root-level orchestrator: runs ingest then build, prints deploy instructions

## data.json schema

```
meta: { generated, currentMonth, notMappedMhb, notMappedRows, badDates, years, plants }
locs: [[lat, lng], ...]
customers: [name, ...]
jobs: [jobCode, project, custIdx, contractor, city, state, zip, [[part,n],...top15], [minYr, maxYr], originalState]
mhb: [locIdx, year, month, plantIdx, qty, jobIdx, prec, [[diam, count], ...]]
```

`jobs[9]` (`originalState`) is `null` for normal records; a 2-letter state code (e.g. `"AZ"`) when `build_map_data.py` auto-corrected the state. The popup renders a green "State corrected" note when this field is set.

`month` (1–12) is included so the frontend can compute YTD comparisons (cap prior year at the same month as `currentMonth`).

## Key data rules

- Rows with Quantity = 0 or blank are excluded
- Valid plants: "Boulder City", "Sulphur Springs", "Plant City"
- Geocoding precision levels:
  - prec 0 = zip centroid (most accurate)
  - prec 1 = city centroid (fuzzy match cutoff 0.85, same state)
  - prec 2 = state centroid (no city/zip data)
  - prec 3 = suspect city/state (city exists in GeoNames but not in listed state, and no zip guard confirms the listed state). Shown as purple pins, hidden by default in the UI.
- Auto-correction: if a suspect record's city exists in exactly **one** GeoNames state and no zip contradicts it, the state is silently corrected before geocoding. The original state is stored in `jobs[9]` and shown in the popup.
- MHB baseline total: 16,839 as of 2026-07-02 (build prints PASS/FAIL against `MHB_BASELINE` dict)
- `generated` date is set from `date.today()` at build time — no need to hardcode

## Quarterly update workflow

1. Download from NetSuite:
   - "AppxShipped" saved search → `../Combined-Database/data/AppxShipped*.xls`
   - "OSLocationData" saved search → `../Combined-Database/data/OSLocationData*.xls`
2. From project root: `python run_quarterly_update.py`
   - Step 1: `ingest_netsuite.py` — appends new rows to NSAW xlsx (auto-detects cutoff)
   - Step 2: `build_map_data.py` — geocodes and writes `output/data.json`
3. Deploy: `cd output && git add data.json index.html && git commit -m "Q refresh" && git push`

Or to skip ingest and just rebuild the map: `python run_quarterly_update.py --skip-ingest`

## Growth KPI (YTD comparison)

The "YTD growth" KPI tile compares the current year's shipments (Jan–`currentMonth`) to the same window of the prior year. This avoids the mid-year dip that appears when comparing a partial year to a full prior year.

Implementation: `month` is stored on each marker's `.meta`; `updateKPI` sums `visible` markers where `year === maxYr - 1 && month <= DATA.meta.currentMonth` as the denominator.

## Input paths

- NSAW xlsx: `C:\Users\JohnLeitzke\OneDrive - Armorock LLC\Documents\Desktop\NSAW All Shipping Data1.2.xlsx` (Sheet1)
- jobcode_db: `..\Schooleys Shit\Jacks_Data_Improvement_Plans\output\jobcode_db.json`
- GeoNames: `data/US.txt`
- NetSuite XLS exports: `..\Combined-Database\data\AppxShipped*.xls` and `OSLocationData*.xls`

## geo_anomalies.csv

Written to `output/geo_anomalies.csv` by both `ingest_netsuite.py` (new records only) and `build_map_data.py` (full historical dataset). Contains only **unresolvable** suspect records (city in multiple states, or zip confirms the listed state). Auto-corrected records do not appear here.

Columns: `job_code, city, listed_state, valid_states, zip, year/date, part_type`

Review workflow: open CSV → correct state in NSAW xlsx for confirmed errors → re-run `build_map_data.py`.
