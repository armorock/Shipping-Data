# Jacks_Data_Improvement_Plans

**Author:** John Leitzke  
**Date:** 2026-05-25  
**Purpose:** Source-traceable job code database with plant conflict detection

---

## What This Is

A standalone data layer that runs independently of your existing scripts. Nothing in here touches `01_build_job_code_registry.py`, `02_jobcode_repair.py`, or any of the `Job Codes/` markdown files. You can keep working on those exactly as before.

The two scripts here read directly from `MASTER CSV FILES/` and your existing `Job Codes/` markdown files, then produce:

1. `output/jobcode_db.json` — one record per job code, every field tagged with its source, conflicts stored inline
2. `output/plant_conflict_report.xlsx` — all job codes where plant data doesn't agree across sources
3. `output/problem_children.xlsx` — the short list that needs manual review

---

## How to Use the JSON in Excel (Power Query)

1. Open Excel → **Data** tab → **Get Data** → **From File** → **From JSON**
2. Point it at `Jacks_Data_Improvement_Plans/output/jobcode_db.json`
3. Power Query opens — click **To Table** in the top left, then **OK**
4. Expand the columns you want by clicking the double-arrow icon on any column header
5. **Close & Load** → build any pivot table from there

Refresh the connection anytime the JSON is regenerated (Data → Refresh All).

---

## How to Run

```
cd Jacks_Data_Improvement_Plans
python jl_build_jobcode_db.py    # step 1 — generates output/jobcode_db.json
python jl_plant_audit.py         # step 2 — generates the two Excel reports
```

Requires: Python 3.8+, openpyxl (`pip install openpyxl`). No other new packages.

---

## What the JSON Contains

One flat object per job code. Every data field has a companion `_source` field showing where the value came from. When sources disagree, a `_conflict` field explains what disagreed and a `_resolution` field says how it was handled.

**Fields:**

| Field | Description |
| --- | --- |
| `job_code` | Primary key — 3-character code |
| `project_name` | Official name from BOM |
| `project_name_source` | Which source provided the name |
| `shipping_city` | Best-available ship-to city |
| `shipping_city_source` | Source of city value |
| `shipping_city_resolution` | MATCH / CONFLICT / SINGLE_SOURCE / CONSENSUS_OVERRIDE |
| `shipping_city_conflict` | Human-readable description if there's a conflict |
| `shipping_state` | Best-available ship-to state (2-letter abbrev) |
| `shipping_zip` | Zip code |
| `shipping_county` | County (from ERP shipping) |
| `customer` | Invoice customer from ERP |
| `contractor` | Contractor name from BOM release document |
| `plant` | Best-available manufacturing plant (BC / SS / PC) |
| `plant_source` | Which source determined the plant value |
| `plant_alec` | What you wrote in the markdown file |
| `plant_resolution` | How the plant value was decided |
| `plant_conflict` | Description of what disagreed (if anything) |
| `year_released` | Year from BOM |
| `date_released` | BOM release date |
| `in_bom` | True/False |
| `in_dispatch` | True/False |
| `in_shipping` | True/False |
| `in_markdown` | True/False |
| `bom_row_count` | Raw BOM line items |
| `dispatch_row_count` | Dispatch board entries |
| `shipping_row_count` | ERP shipping line items |

---

## Conflict Resolution Logic

There are 4 independent data sources. Shipping (QB, FB, NS) counts as **one** source — all three are the same ERP pipeline and share errors, so QB + FB + NS agreeing is not 3 independent votes.

| Group | Source | Confidence |
| --- | --- | --- |
| A | BOM (release documents) | Highest |
| B | Dispatch Board | High |
| C | Shipping / ERP (QB + FB + NS combined) | Medium |
| D | Markdown files (manually maintained) | Lower |

**Rules applied per field:**

- `MATCH` — all sources agree
- `CONSENSUS_OVERRIDE` — BOM says X, but Dispatch + Shipping + Markdown all independently say Y → use Y, flag for review
- `CONFLICT` — sources disagree and no clean consensus exists → keep highest-confidence value, flag for review
- `SINGLE_SOURCE` — only one source has this field, no conflict possible

For `plant` specifically: BOM rarely carries plant data, so resolution is between Dispatch, Shipping, and Markdown. Dispatch is treated as highest-confidence for plant.

---

## Known Issues in the Existing Scripts (Your Call Whether to Fix)

These are observations John found while building this — Alec decides if/when to apply them:

1. **`02_jobcode_repair.py` line 282** — opens the BABY file with `keep_vba=True` instead of `data_only=True`. Excel VLOOKUP formulas in the state/city columns are read as literal formula strings (e.g., `=IFERROR(VLOOKUP(...))`) instead of their evaluated values. This is why approximately 3,354 rows show as UNRESOLVABLE.

2. **`01_build_job_code_registry.py`** — the BOM has a `Contractor` column that is never extracted into the registry. Phase 2 can't use it for matching even though contractor name is a strong signal.

3. **`02_jobcode_repair.py` `score_candidate()`** — City is read from both sources and passed into the scoring function but is never actually scored. Adding it as a criterion would improve match precision.

---

## File Structure

```
Jacks_Data_Improvement_Plans/
├── README.md                     <- this file
├── jl_build_jobcode_db.py        <- run first
├── jl_plant_audit.py             <- run second
└── output/
    ├── jobcode_db.json           <- Power Query connects here
    ├── plant_conflict_report.xlsx
    └── problem_children.xlsx
```
