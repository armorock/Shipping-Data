# Shipping-Data

Root repository for Armorock's shipping data pipeline and map dashboard. Contains two active sub-projects and a root-level orchestrator.

## Sub-projects

| Folder | Purpose |
|---|---|
| `Shipping-Map/` | Leaflet dashboard of MHB bases shipped (2014–present). Live on GitHub Pages. |
| `Combined-Database/` | Six-module pipeline combining BOM, Dispatch, and ERP sources into an auditable source of truth. |
| `Extracting-BOM-SharePoint/` | Pulls BOM documents from SharePoint for Combined-Database input. |
| `Extracting-M-Drive/` | Pulls New Customer Forms and shop drawings from the M: drive. |
| `BOM-Structure-Detail/` | Detailed BOM extraction per structure. |
| `Schooleys Shit/Jacks_Data_Improvement_Plans/` | Produces `jobcode_db.json` used by the map for location enrichment. |

## Quarterly update (entry point)

```sh
# Prerequisites (manual — require network access):
#   Download NetSuite exports into Combined-Database/data/
#   Run BOM SharePoint / M-drive extraction if source data changed

python run_quarterly_update.py
```

This runs:
1. `Shipping-Map/ingest_netsuite.py` — appends new NetSuite records to NSAW xlsx
2. `Shipping-Map/build_map_data.py` — geocodes, writes data.json + index.html

Then deploy the Pages site:
```sh
cd Shipping-Map/output
git add data.json index.html && git commit -m "Q refresh" && git push
```

Flag: `--skip-ingest` to skip step 1 and just rebuild the map.

## Key files

- `run_quarterly_update.py` — orchestrator (see above)
- `Shipping-Map/ingest_netsuite.py` — NetSuite XLS ingest; auto-detects NSAW cutoff date
- `Shipping-Map/build_map_data.py` — geocoder and data.json builder
- `Shipping-Map/index.html` — map dashboard source (copied to `Shipping-Map/output/` on build)
- `Shipping-Map/output/` — standalone git repo pushed to `armorock/ax-fieldatlas-7q3x` (GitHub Pages)
