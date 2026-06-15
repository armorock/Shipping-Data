# Combined-Database — Armorock Unified Source of Truth

Builds a clean, auditable source of truth for Armorock structures/shipping from the upstream
primary sources (BOM union, Dispatch Board, ERP/NetSuite shipping, the job-code registry), with
full provenance, trust-based conflict resolution learned from history, deduplication so no physical
piece is double-counted, lifecycle separation (shipped vs quoted-not-shipped), and quarterly
re-runnability with change tracking.

The old `NSAW All Shipping Data1.2.xlsx` workbook is **deprecated** — it was a derived mashup, not a
source. This project replaces it.

## How to run

```sh
python run_quarter.py --quarter 2026-Q2     # full pipeline + snapshot/diff
```

Or step by step:

```sh
python trust_model.py            # learn field x source trust -> output/trust_report.md
python build_observation_db.py   # every reported value -> output/observations.db
python build_registry.py         # Layer 1 -> output/project_registry.csv
python build_pieces_ledger.py    # Layer 2 -> output/pieces_ledger.csv (+ reconciliation, quoted_not_shipped)
python build_reports.py          # rehab/nonbase, provenance, completeness roadmap, review lists
python snapshot_and_diff.py 2026-Q2
```

## Review round-trip (fixing data)

1. Open `output/review/` (see `_index.md`). Each list is independent.
2. Fill the **Confirmed Value** column on rows you want to fix.
3. `python apply_corrections.py 02_city.xlsx` (or no args = all lists).

Confirmed choices persist in `data/resolutions.csv` and are honored on every future run, so review
effort is cumulative — you never re-decide the same gap.

## The model

- **Project Registry** (`project_registry.csv`) — one row per `job_code`: resolved
  location/customer/contractor/plant/year with conflict status and a `needs_review` flag.
- **Pieces Ledger** (`pieces_ledger.csv`) — one row per physical piece shipped, classified
  `base / rehab / non-base / unknown`, with location denormalized from the registry.

## Key rules

- **One job = one city** (~99%): a job_code resolves to a single jobsite.
- **Shipped-count overlap** (`SHIPPED_OVERLAP` in `sources.py`): for 2019–2025 both Dispatch and ERP
  record shipments. The reconciliation showed Dispatch is materially incomplete (~21.5k vs ~66.5k
  pieces, adds only 51 unique jobs), so the default is **`union`** = ERP all years + pieces from
  jobs that appear only in Dispatch (no double-count). See `output/dispatch_vs_erp_reconciliation.xlsx`.
- **Classification**: part number is authoritative (carries the `RMH` rehab marker the human
  `part_type` column often mislabels as "section"), then the human word, across all four naming
  generations (Gen1–Gen4).
- **No phantom/duplicate pieces**: ERP QB/FB/NetSuite triplicates are deduped (max per system);
  only rows with a part number count; quoted/released-not-shipped never count as shipped.

## Inputs (read-only, from existing projects — see `sources.py`)

- `Schooleys Shit/MASTER CSV FILES/all_bom_union.xlsx` — released/planned layer (BOM 2016–2026)
- `Schooleys Shit/MASTER CSV FILES/Dispatch_Board_Master_2019-2025.csv` — clean shipped ledger
- `Schooleys Shit/MASTER CSV FILES/All Shipping Data BABY.xlsm` (Master List) — ERP shipping ledger
- `Schooleys Shit/Jacks_Data_Improvement_Plans/output/jobcode_db.json` — registry basis + conflict history
- `Schooleys Shit/jobcode_repair_log.csv` — repair history (trust calibration)
- `Shipping-Map/data/US.txt` — GeoNames centroids

## Outputs (`output/`)

`observations.db`, `project_registry.csv`, `pieces_ledger.csv`, `rehab_nonbase_report.xlsx`,
`quoted_not_shipped.xlsx`, `dispatch_vs_erp_reconciliation.xlsx`, `structure_completeness_roadmap.xlsx`,
`project_provenance.xlsx`, `trust_report.md`, `review/`, `corrections_log.csv`,
`snapshots/<quarter>/`, `snapshot_metrics.csv`, `changes_<quarter>.xlsx`.
