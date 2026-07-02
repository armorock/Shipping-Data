# Shipping-Data — root project notes

Multi-sub-project repo. The two active pipelines are `Combined-Database/` and `Shipping-Map/`. Each has its own `CLAUDE.md` with detailed notes. See those first when working in those folders.

## Root-level files

- `run_quarterly_update.py` — orchestrator: validates prerequisites, runs `ingest_netsuite.py` then `build_map_data.py`, prints deploy instructions. Flags: `--skip-ingest`.

## Quarterly update flow

1. **Manual prerequisites** (require network):
   - Download NetSuite "AppxShipped" saved search → `Combined-Database/data/AppxShipped*.xls`
   - Download NetSuite "OSLocationData" saved search → `Combined-Database/data/OSLocationData*.xls`
   - BOM SharePoint extraction and M-drive extraction if source data changed (run those sub-project scripts separately)

2. `python run_quarterly_update.py` from this root

3. Deploy: `cd Shipping-Map/output && git add data.json index.html && git commit -m "Q refresh" && git push`

## Sub-project CLAUDE.md locations

- `Shipping-Map/CLAUDE.md` — map dashboard, data.json schema, geocoding rules, geo anomaly detection
- `Combined-Database/CLAUDE.md` — six-module pipeline architecture, column schemas, scale stats
