# CLOCK Code Audit — Schooley's Shipping Data

A Python-based audit pipeline that cross-references BOM, Dispatch Board, and ERP Shipping records to populate the Job Code Master List and surface every meaningful discrepancy across Armorock's full production history (2014–2026).

## What It Does

For every job code, the pipeline answers: *what was released to production, what was dispatched, what actually shipped, and where are the gaps?*

See [CLOCK_AUDIT_MASTER_PLAN.md](CLOCK_AUDIT_MASTER_PLAN.md) for the full 9-phase execution plan.

## Data Sources

| Source | File | Rows | Job Codes | Date Range |
|---|---|---|---|---|
| BOM 2016–2026 | `MASTER CSV FILES/all_bom_union.xlsx` | 216,005 | 2,104 | 2016–2026 |
| Dispatch Board | `MASTER CSV FILES/Dispatch_Board_Master_2019-2025.csv` | 27,353 | 1,924 | 2019–2025 |
| ERP Shipping | `MASTER CSV FILES/All Shipping Data BABY.xlsm` | 93,653 | 2,492 | 2014–2026 |
| Job Code Universe | `MASTER CSV FILES/Job Code Master List.xlsx` | 3,000 | 3,000 | — |

Raw dispatch board source files are in `Raw information from dispatch boards/`.

## How to Run

Scripts run in phase order. Each phase depends on the previous.

```
python 01_build_job_code_registry.py   # Phase 1 — Job Code Registry
python 02_jobcode_repair.py            # Phase 2 — BABY file job code repair
# ⏸ PAUSE — review repair output before continuing
python generate_master_list.py
python generate_job_codes.py
python generate_dashboards.py
```

## Outputs

- `Job_Code_Registry.xlsx` — one row per job code, all sources aggregated
- `Job Codes/` — per-job-code markdown files (one per job code, ~3,000 files)
- `Cities/` — per-city markdown rollups
- `States/` — per-state markdown rollups
- `jobcode_repair_log.csv` — Phase 2 repair decisions

## Current Status

Phase 1 (Job Code Registry) is built and running. Phase 2 (BABY file job code repair) is implemented. See `CLOCK_AUDIT_MASTER_PLAN.md` for remaining phases 3–9.
