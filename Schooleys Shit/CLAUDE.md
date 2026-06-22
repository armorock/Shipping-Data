# Schooleys Shit — Project Context for Claude Code

## What This Project Is

A Python audit pipeline that cross-references four data sources (BOM, Dispatch Board, ERP Shipping,
and Job Code markdown files) to populate a Job Code Master List and surface every meaningful
discrepancy. Goal: for every job code, answer what was released, what was dispatched, what shipped,
and where are the gaps.

**Two parallel workstreams live here — they are fully independent:**

| Workstream | Owner | Scripts | Status |
|---|---|---|---|
| CLOCK Audit Pipeline | Alec Schooley | `01_build_job_code_registry.py`, `02_jobcode_repair.py` | Phase 2 / PAUSE & REVIEW |
| Job Code DB + Plant Audit | John Leitzke | `Jacks_Data_Improvement_Plans/` | Complete — outputs in `output/` |

Neither workstream touches the other's files.

---

## Data Sources

| Source | File | Rows | Job Codes | Date Range |
|---|---|---|---|---|
| BOM 2016–2026 | `MASTER CSV FILES/all_bom_union.xlsx` → sheet `BOM 2016-2026` | 462,552 | 2,304 | 2016–2026 |
| Dispatch Board | `MASTER CSV FILES/Dispatch_Board_Master_2019-2025.csv` | 27,353 | 1,924 | 2019–2025 |
| ERP Shipping | `MASTER CSV FILES/All Shipping Data BABY.xlsm` → sheet `Master List` | 93,653 | 2,492 | 2014–2026 |
| Job Code markdown files | `Job Codes/*.md` | ~3,000 files | ~3,000 | — |

Job codes are 3-character alphanumeric, first letter A–E, pattern `^[A-E][A-Z0-9]{2}`.

Plant codes: `BC` = Boulder City, `SS` = Sulphur Springs, `PC` = Plant City.

ERP eras in shipping data: QB = QuickBooks (2014–2023), FB = Fishbowl, NS = NetSuite.

---

## Alec's Pipeline (CLOCK Audit)

### Phase 1 — `01_build_job_code_registry.py`
Reads `all_bom_union.xlsx` and the `Job Codes/*.md` markdown files. Produces `Job_Code_Registry.xlsx`.

**Known gap:** The BOM has a `Contractor` column that is never extracted. This matters for Phase 2
matching because contractor name is a strong signal.

### Phase 2 — `02_jobcode_repair.py`
Reads `Job_Code_Registry.xlsx` and `All Shipping Data BABY.xlsm`. Attempts to repair ~3,354
UNRESOLVABLE job code rows by matching against BOM records. Produces `jobcode_repair_log.csv` and
`02_repair_output.xlsx`.

**Known bug — line 282:** Opens the BABY file with `keep_vba=True` instead of `data_only=True`.
This causes openpyxl to read Excel VLOOKUP formula strings (e.g., `=IFERROR(VLOOKUP(...))`) instead
of their evaluated values — which is why ~3,354 rows are UNRESOLVABLE. Fix: change line 282 to
`openpyxl.load_workbook(BABY_PATH, read_only=True, data_only=True)`.

**Known gap — `score_candidate()`:** City is extracted from both sources and passed into the scoring
function but is never actually scored. Adding city as a criterion would improve match precision.

Current scoring: Prefix (1pt) + Customer (1pt) + State (1pt) + Date ±180 days (1pt).
Thresholds: ≥3 = High confidence, 2 = Medium.

---

## John's Layer — `Jacks_Data_Improvement_Plans/`

Built entirely independently. Does not modify any of Alec's files or outputs.

### What it does

Reads the same raw source files directly and produces:
- `output/jobcode_db.json` — one flat record per job code, every field tagged with its source,
  conflicts stored inline. Designed to be loaded directly by Power Query (no expansion steps needed).
- `output/plant_conflict_report.xlsx` — all job codes where plant data doesn't fully agree
- `output/problem_children.xlsx` — the short list (CONFLICT + CONSENSUS_OVERRIDE only) requiring
  manual review

### Source independence model

There are 4 independent sources. Shipping (QB + FB + NS) counts as **one** source — they share the
same ERP pipeline and errors, so three ERP systems agreeing is not three independent votes.

| Group | Source | Confidence |
|---|---|---|
| A | BOM (release documents) | Highest |
| B | Dispatch Board | High |
| C | Shipping / ERP (QB+FB+NS combined) | Medium |
| D | Markdown files | Lower |

### Conflict resolution

- `MATCH` — all sources agree
- `CONSENSUS_OVERRIDE` — BOM says X, but Dispatch + Shipping + Markdown all say Y → use Y, flag
- `CONFLICT` — sources disagree, no clean consensus → keep highest-confidence value, flag
- `SINGLE_SOURCE` — only one source has this field
- `NO_DATA` — no source has this field

For plant specifically: BOM rarely carries plant data, so resolution is between Dispatch (highest
confidence for plant), Shipping, and Markdown.

### JSON schema (key fields per record)

```
job_code, project_name, project_name_source
shipping_city, shipping_city_source, shipping_city_resolution, shipping_city_conflict
shipping_state, shipping_state_source, shipping_state_resolution, shipping_state_conflict
shipping_zip, shipping_county, customer, customer_source, contractor, contractor_source
plant, plant_source, plant_alec, plant_resolution, plant_conflict
year_released, date_released
in_bom, in_dispatch, in_shipping, in_markdown
bom_row_count, dispatch_row_count, shipping_row_count
```

### Results from last run (2026-06-22)

3,071 total records. Plant resolution breakdown:
- MATCH: 1,743
- SINGLE_SOURCE: 249
- CONFLICT: 37
- CONSENSUS_OVERRIDE: 0
- NO_DATA: 639

### `jl_sharepoint_reader.py`

Reads raw documents directly from the 4 SharePoint year sites (2023–2026) using the Microsoft
Graph API. Parses BOM XML/PDF, NCF/ECF docx, and Shop Drawing PDFs per job folder and produces
`output/jobcode_db_sharepoint.json` with cross-document comparison fields:
- `location_bom` vs `location_ncf` + `location_conflict`
- `contractor_bom` vs `customer_ncf` + `contractor_customer_match`
- `structure_count_bom` vs `structure_count_shop_drawing` + `structure_count_conflict`
- `documents_found` / `documents_missing` per job

Requires auth via Microsoft Graph (device-code OAuth, token cached at `~/.claude/msgraph_token.json`).
Infrastructure lives in `../BOM-Structure-Detail/` (graph_client.py, sharepoint_client.py,
parse_bom_pdf.py) — the script adds that path to sys.path at startup.

### `jl_mdrive_reader.py`

Combines M: Drive (2016–2022, local network) and SharePoint (2023–2026, Graph API) into a single
pass, producing `output/jobcode_db_mdrive.json`. Intended as an 8th source for `jl_build_jobcode_db.py`
(integration not yet wired). Year priority: oldest processed first, newest year wins on collision.

Same output schema as `jobcode_db_sharepoint.json`. Key differences:

- M: Drive years require a `_Release` subfolder inside each job folder — jobs without one produce
  an empty record (planned fix: fall back to scanning the full job folder)
- NCF parsing disabled for all M: Drive years and for SharePoint 2023
- Contractor extracted from PO filename on M: Drive (no download needed)
- Early years (2016–2018) often have only QUOTE files — no BOM → empty location/contractor
  (planned fix: parse quotes as fallback)

Requires the same Graph auth as `jl_sharepoint_reader.py`.

### `parse_bom_pdf.py` — structure name extraction fix (2026-05-26)

Applied to both `BOM-Structure-Detail/parse_bom_pdf.py` and `Extracting-M-Drive/parse_bom_pdf.py`.

`_find_structure_name` and `parse_shop_drawing_pdf` previously took the entire pdfplumber layout
line as the structure name. With `layout=True`, two-column shop drawing headers merge into one line:

```
PO Box 60006               Structure:   SSMH #1
```

The fix: when the pattern match is preceded by a colon (a label), extract from the match position
only. When the line IS the structure name (no preceding colon), take the whole line. This gives
clean names like "SSMH #1" that match NetSuite and other systems.

### How to run

```
cd Jacks_Data_Improvement_Plans
python jl_build_jobcode_db.py       # step 1 — builds output/jobcode_db.json
python jl_plant_audit.py            # step 2 — builds the two Excel reports
python jl_sharepoint_reader.py      # optional — builds output/jobcode_db_sharepoint.json
python jl_mdrive_reader.py          # optional — builds output/jobcode_db_mdrive.json (M: Drive + SP)
```

### How to use the JSON in Excel (Power Query)

Data tab → Get Data → From File → From JSON → point at `output/jobcode_db.json` → To Table → OK
→ expand columns → Close & Load → build pivot tables. Refresh anytime the JSON is regenerated.

---

## File Structure

```
Schooleys Shit/
├── CLAUDE.md                          ← you are here
├── CLOCK_AUDIT_MASTER_PLAN.md         ← Alec's full execution plan
├── README.md                          ← project overview
├── 01_build_job_code_registry.py      ← Alec's Phase 1
├── 02_jobcode_repair.py               ← Alec's Phase 2
├── Job_Code_Registry.xlsx             ← Phase 1 output
├── jobcode_repair_log.csv             ← Phase 2 output
├── 02_repair_output.xlsx              ← Phase 2 output
├── Job Codes/                         ← ~3,000 markdown files (one per job code)
├── MASTER CSV FILES/                  ← shared raw source data (read-only)
│   ├── all_bom_union.xlsx
│   ├── Dispatch_Board_Master_2019-2025.csv
│   └── All Shipping Data BABY.xlsm
└── Jacks_Data_Improvement_Plans/
    ├── README.md                      ← explains John's layer for humans
    ├── jl_build_jobcode_db.py
    ├── jl_plant_audit.py
    ├── jl_sharepoint_reader.py
    └── output/
        ├── jobcode_db.json            ← Power Query connects here
        ├── jobcode_db_sharepoint.json
        ├── plant_conflict_report.xlsx
        └── problem_children.xlsx
```
