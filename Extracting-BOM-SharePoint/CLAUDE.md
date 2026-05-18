# CLAUDE.md — Project Instructions

## Project Purpose

Extract Armorock BOM PDFs/XMLs from SharePoint and produce structured CSVs to support a **map of every structure Armorock has shipped, by location**. Each row in the output represents one part in one structure on one job. The target output format is defined in `data/Example Formatting.csv` and `data/ex.bom.EHP.csv`.

## End Goal: Structure Shipping Map

The ultimate deliverable is a map showing where every manhole/structure has been shipped, with drill-down into what parts comprised each structure.

### Location data available

| Field | Source | Notes |
|---|---|---|
| `Job Location` | BOM header or NCF `Job Site City/State` | City, State (e.g. "Denver, CO") — sufficient for job-level map pins |
| `Zip Code` | NCF `Job Site Zip Code` | Only populated when Job Site City+State confirmed; blank for pre-2024 |
| `Structure Name` | BOM page header | Station reference (e.g. `MH 01 STA 00+00.00`) — NOT GPS coordinates |

### Data model rationale

- The output is **one row per line item per structure** (flat/long format). This is intentional — it is the data layer.
- A future "one row per structure" aggregated view (for the map layer) should be built on top of this CSV, not replace it.
- **Include all categories** (Precast AND Resale items like Joint Seal, Hardware, Frame & Ring).
- Job-level fields (`Year Release`, `Job Code`, `Project Name`, `Job Location`, `Contractor`) repeat on every row — normal for flat CSVs.

## File Layout

```
data/                          Source files: example PDFs, reference CSVs, sample BOMs
data/state_abbreviations.csv  Full state name → 2-letter abbreviation lookup (all 50 states + DC + territories)
data/location_overrides.csv   Manual Job Code → Job Location overrides for jobs with no BOM/NCF location
output/                        All script outputs go here — never write to project root
```

## SharePoint Configuration

| Year | SITE_PATH | DRIVE_NAME | Script |
| --- | --- | --- | --- |
| 2026 | `/sites/JobData2026` | `Job Data 2026` | `extract_bom_test.py` |
| 2025 | `/sites/jobdata2025` | `Job Data 2025` | `extract_bom_2025.py` |
| 2024 | `/sites/jobdata2024` | `Job Data 2024` | `extract_bom_2024.py` |
| 2023 | `/sites/jobdata2023` | `Job Data 2023` | `extract_bom_2023.py` |

- Job folders at drive root follow pattern `XXX - Name` (e.g. `ELC - ALCOSAN Wet Weather PS`)
- BOM files in subfolders (`5-Release/` etc.) — found by `iter_files(recursive=True)`
- Auth: `~/.claude/msgraph_config.json` (`tenant_id`, `client_id`), device-code flow, token cached at `~/.claude/msgraph_token.json`

## Document Hierarchy

When extracting data, use this priority order (prefer most detailed source):

1. **BOM XML** (`*-BOM by Structure (Excel XML).xml`) — preferred, no PDF crash risk
2. **BOM by Structure PDF** — one page per structure, structure name from page header
3. **BOM Summary PDF** — fallback, structure name parsed from filename
4. **NCF/ECF `.docx`** (`_NCF.docx` / `_ECF.docx`) — location fallback only; trusted for 2024+ only
   - Provides: `Job Site City/State` → `Job Location`, `Job Site Zip Code`, contractor name
   - If Job Site City/State blank, falls back to Billing City/State (zip stays blank in that case)
   - If no `_NCF`/`_ECF` file found, tries any `.docx` in the folder

## File Roles

| File | Role |
|---|---|
| `graph_client.py` | Auth token, `graph_get`, `graph_get_all` (pagination); `ensure_fresh_token()` proactive refresh with 5-min buffer and Windows chime; `expires_at` stored in token cache |
| `sharepoint_client.py` | `get_site`, `get_drive`, `list_children`, `iter_files`, `download_file` |
| `parse_bom_pdf.py` | `parse_bom_pdf`, `parse_bom_by_structure_pdf`, `parse_bom_by_structure_xml`, `parse_ncf_docx`, `normalize_location`; subprocess-safe wrappers `parse_bom_pdf_safe`, `parse_bom_by_structure_pdf_safe`; loads `data/state_abbreviations.csv` at import time |
| `extract_bom_test.py` | 2026 extraction → `output/bom_manhole_map.csv` |
| `extract_bom_2025.py` | 2025 extraction → `output/bom_manhole_map_2025.csv` |
| `extract_bom_2024.py` | 2024 extraction → `output/bom_manhole_map_2024.csv` |
| `extract_bom_2023.py` | 2023 extraction → `output/bom_manhole_map_2023.csv` (no NCF) |

## Output Column Schema

| Column | Source |
|---|---|
| `Year Release` | Date in filename (`4.20.26` → `2026`) or drive year |
| `BOM Release Date` | `Reported On` field in XML header, or release date from PDF |
| `Date extracted` | Today's date when the script runs |
| `Job Code` | 2–4 letter prefix from job folder name |
| `Project Name` | `Job Name` from BOM header; NCF `Name of Job` as fallback |
| `Structure Name` | Page header in BOM by Structure; filename suffix for BOM Summary |
| `Job Location` | `Location` from BOM header; NCF `Job Site City, State` as fallback |
| `Location Source` | `BOM` / `NCF Job Site` / `NCF Billing` / `Manual` / blank — always indicates where location came from |
| `Zip Code` | NCF `Job Site Zip Code` (2024+ only; blank if job site city/state not confirmed) |
| `Contractor` | `Contractor` from BOM header; NCF customer as fallback |
| `Agency` | Blank — comes from New Customer Form (not yet parsed separately) |
| `Engineer` | Blank — comes from New Customer Form (not yet parsed separately) |
| `Part Name` | Decoded from part number via `build_part_name`; falls back to BOM description |
| `Product Number` | Line item part number |
| `Quantity` | Line item quantity |
| `Weight` | Weight (lbs); blank for non-precast |
| `Production Part` | BOM category (Precast, Joint Seal, Frame & Ring, Hardware, etc.) |
| `Part Type` | For Precast: from part number prefix (MHL=Lid, MHS=Section, MHB=Base, MHC=Cone, MHGR=Grade Ring). For others: "Resale" |
| `Part Subtype` | For Precast: ECC=Eccentric, FLAT=Flat, etc. For Joint Seal: Mastic, Gasket, etc. |

## BOM Filename Conventions

Pattern: `{CODE} {M.D.YY} {Job Name}-{Structure}-BOM by Structure.pdf`
or:      `{CODE} {M.D.YY} {Job Name}-{Structure}-BOM Summary.pdf`
or:      `{CODE} {M.D.YY} {Job Name}-BOM by Structure (Excel XML).xml`

## Folder Regex

`^([A-Z]{2,4})\s*[-–—û]\s*(.+)$` — matches `CVG- Name`, `ELC - Name`, and folders using Unicode dash variants including `û`.

## State Name Normalization

All `Job Location` values are passed through `normalize_location()` before output. This converts full state names (e.g. "Florida", "North Carolina") to 2-letter abbreviations ("FL", "NC") so BOM-sourced and NCF-sourced locations are consistent.

- Source of truth: `data/state_abbreviations.csv` — edit this file to add or correct mappings
- Applied in `build_row` in all four extract scripts
- If the CSV is missing, normalization silently passes the value through unchanged

## Location Overrides

`data/location_overrides.csv` provides manual locations for jobs where BOM and NCF both return blank. Columns: `Job Code`, `Job Location`, `Location Source` (use `"Manual"`). Applied as a last resort in `build_row` across all four scripts. Add rows here whenever a job location is confirmed but not in any source document.

## PDF Crash Isolation

`pdfplumber` can cause native C-level crashes that kill the process. All PDF parsing goes through subprocess wrappers (`parse_bom_pdf_safe`, `parse_bom_by_structure_pdf_safe`) that run the parser in a child process and return results via pickle. XML BOMs are parsed in-process (no crash risk).

## Sync Rules

This project receives **structural/search changes only** from BOM Structure Detail — no data-pull changes:
- `_STRUCTURE_PATTERNS` and `_NON_STRUCTURE_PATTERNS` stay in sync with BOM Structure Detail and M drive
- `graph_client.py` stays in sync with BOM Structure Detail (token management, retry logic)
- Polymer classification (`PR*`/`HDPE-*` → `Part Type = "Polymer"`) is **not** applied here — output schema is unchanged
- New output columns (errors.csv, unclassified_files.csv) are **not** applied here

## GitHub

Repository: `https://github.com/armorock/bom-sharepoint` (private)

## Coding Conventions

- No explanatory comments
- No premature abstractions
- Validate only at system boundaries
- All outputs to `output/` — use `os.makedirs("output", exist_ok=True)` before writing
