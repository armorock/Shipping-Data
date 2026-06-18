# CLAUDE.md — BOM Structure Detail

## Project Purpose

Extract **every BOM file** from every job folder across all four year SharePoints and produce
per-year CSVs for cross-referencing against accounting data (NetSuite).  This is the
unfiltered version of the original extraction — no file-type priority, no structure-name
deduplication.  The goal is a complete record of everything ever released so accounting can
be used as the source of truth to confirm what was actually shipped.

## Key Differences from Original Extraction

| Behavior | Original (`extract_bom_*.py`) | This project (`extract_all.py`) |
|---|---|---|
| File type priority | XML beats PDF; PDF beats Summary | ALL types processed per job |
| Duplicate structures | `seen_structures` skips if name already seen | No deduplication — same name from two files = two sets of rows |
| Add-on/remake subfolders | Picked up only if structure name differs | Always included, tagged with `Source Subfolder` |
| Source tracking | None | `Source File`, `Source Subfolder`, `BOM Type` columns on every row |
| Non-standard folder names | Skipped entirely | Processed; job code inferred via 4-step resolution |
| Token expiry mid-run | 401 causes FOLDER ERROR, skips remaining folders | Proactive refresh before each folder via `ensure_fresh_token()`; plays Windows chime on refresh |

## Observed Folder Structure (from live SharePoint browse, 2026)

Every job folder follows this pattern:
```
{CODE} - {Project Name}/
  1-Plans/
  2-Quote/
  3-Submittals/
  4-Approvals/
  5-Release/               ← BOM files live here
    {NCF}.docx
    {CODE} {M.D.YY} {Name}-BOM by Structure (Excel XML).xml
    {CODE} {M.D.YY} {Name}-BOM by Structure.pdf
    {CODE} {M.D.YY} {Name}-BOM Summary.pdf
    {CODE} {M.D.YY} {Name}-NetSuite CSV.csv
    AR Files/              ← AR copies, prelien notices — no BOM data
    {Remake/Add-on subfolder}/    ← additional releases (see below)
      {CODE} {M.D.YY} {Name}-{Tag}-BOM by Structure (Excel XML).xml
      {CODE} {M.D.YY} {Name}-{Tag}-NetSuite CSV.csv
      ...
    Old reports for reference only/   ← superseded releases; NOT skipped by iter_files
  6-Shipping/
  Antiquated Files/        ← skipped by SKIP_FOLDERS
```

### Add-on / Remake subfolders (confirmed examples)

- `EGQ/5-Release/Base Remake 4.29.26/` — remake of one structure after original release
- `EGR/5-Release/Add-on Grout - 1.14.25/` — add-on purchase order (no BOM file, just quotation)
- `EGR/5-Release/Old reports for reference only/` — superseded BOMs; iter_files recurses into these

### NetSuite CSV files

Each release (including remake subfolders) contains a `*-NetSuite CSV.csv`.  These appear
to be release-time snapshots exported from the quoting/ERP system.  They are not the
accounting records — the full accounting data lives elsewhere and is not in this project.

## SharePoint Configuration

| Year | SITE_PATH | DRIVE_NAME | Output |
|---|---|---|---|
| 2026 | `/sites/JobData2026` | `Job Data 2026` | `output/bom_manhole_map.csv` |
| 2025 | `/sites/jobdata2025` | `Job Data 2025` | `output/bom_manhole_map_2025.csv` |
| 2024 | `/sites/jobdata2024` | `Job Data 2024` | `output/bom_manhole_map_2024.csv` |
| 2023 | `/sites/jobdata2023` | `Job Data 2023` | `output/bom_manhole_map_2023.csv` |

Auth: `~/.claude/msgraph_config.json` (`tenant_id`, `client_id`, optional `client_secret`),
device-code flow, token cached at `~/.claude/msgraph_token.json`.
Token expiry stored as `expires_at`; `ensure_fresh_token()` refreshes proactively (>5 min buffer)
before each folder so runs longer than 1 hour do not lose folders to 401 errors.

NCF parsing is disabled for 2023 (`use_ncf = year != "2023"`) because 2023 NCF files are
not in a consistent location or format.

## Usage

```
python extract_all.py 2026          # one year
python extract_all.py 2025 2026     # multiple years
python extract_all.py               # all four years (sequential)
```

To run all four years in parallel, open 4 terminals and run one year each — output files
are separate so there is no collision.  Token auto-refreshes on 401 mid-run, so long runs
no longer lose folders to expiry.  If a full device-code re-auth is needed (refresh token
also expired), the script will print the device-code prompt and wait.

## File Roles

| File | Role |
|---|---|
| `graph_client.py` | Auth token, `graph_get`, `graph_get_all` (pagination), `ensure_fresh_token()` proactive refresh |
| `sharepoint_client.py` | `get_site`, `get_drive`, `list_children`, `download_file` |
| `parse_bom_pdf.py` | All BOM/NCF/Shop Drawing parsers; subprocess-safe PDF wrappers |
| `extract_all.py` | Main extractor — all years, no filtering |
| `union_all.py` | Combines per-year CSVs into `output/all_bom_2023_2026.xlsm`; frozen header, auto-filter, column widths |
| `extract_bom_2026.py` | Single-year 2026 extractor (legacy; prefer `extract_all.py`) |
| `extract_bom_2025.py` | Single-year 2025 extractor (legacy; prefer `extract_all.py`) |
| `extract_bom_2024.py` | Single-year 2024 extractor (legacy; prefer `extract_all.py`) |
| `extract_bom_2023.py` | Single-year 2023 extractor (legacy; prefer `extract_all.py`) |
| `data/document_types.csv` | Lookup table: filename keyword + extension → Source File Name label |

## Output Column Schema

| Column | Notes |
|---|---|
| `Year Release` | Year from BOM release date; falls back to SharePoint drive year |
| `BOM Release Date` | M/D/YY date from BOM header |
| `Date extracted` | Date this script was run |
| `Job Code` | 2–4 uppercase letters, resolved via 4-step process (see below) |
| `Project Name` | From BOM header or NCF |
| `Structure Name` | Per-structure for XML/ByStructure; derived from filename for Summary |
| `Job Location` | City/state from BOM header, NCF, or `data/location_overrides.csv` |
| `Location Source` | `BOM`, `NCF`, or override source label |
| `Zip Code` | From NCF only |
| `Contractor` | From BOM header or NCF customer field |
| `Agency` | Blank (not extracted) |
| `Engineer` | Blank (not extracted) |
| `Part Name` | Human-readable name built from product number; falls back to description |
| `Product Number` | Raw part number from BOM |
| `Quantity` | From BOM line item |
| `Weight` | From BOM line item (lbs) |
| `Production Part` | Category string from BOM (e.g. "Precast", "Resale") |
| `Part Type` | Classified type (e.g. "Base", "Section", "Cone", "Resale") |
| `Part Subtype` | Classified subtype (e.g. "Eccentric", "Flat Floor", "Mastic", "Pro-Ring", "Ladtech") |
| `Source File` | Exact filename of the document this row came from |
| `Source Subfolder` | Relative path within job folder (e.g. `5-Release/Base Remake 4.29.26`); blank if at job root |
| `Source File Name` | Document type label from `data/document_types.csv` (e.g. `BOM by Structure XML`, `Shop Drawing PDF`) |

## Folder Handling

### System folders skipped inside job folders

`SKIP_FOLDERS = {"forms", "plugin_data", "robotinterface", "__macosx", "antiquated files"}`

`Old reports for reference only` is NOT in this list — superseded BOMs are included with
their subfolder path visible.  Filter in post-processing if needed.

### Non-standard top-level folders (e.g. "Misc. Sales - 2025", "Transfers", "_PC", "_BC", "_SS")

All top-level folders are processed regardless of name.  Job code is resolved in this order:

1. **Folder name** — matches `{CODE} - {Name}` or `{CODE}- {Name}` pattern (e.g. `CVA- Canterwood Manholes`)
2. **BOM filename prefix** — first 2–4 uppercase letters of the BOM filename (e.g. `EKQ 5.15.25...xml` → `EKQ`)
3. **Subfolder path first component** — first segment of the Source Subfolder path, matched against the same `{CODE}-` pattern (handles category folders like `_PC` that contain real job subfolders, e.g. `_PC / CVA- Canterwood Manholes / _release /` → `CVA`)
4. **BOM content header** — `Job Number` then `Job Name` fields from a parsed BOM, only when structures are present; retroactively updates any rows already written for that job

If no BOM files exist in the folder, it prints "No BOM files found — skipping" and moves on harmlessly.

### Valid job code format

All valid job codes start with A, B, C, D, or E (the company has not yet reached F).  Job
codes not starting with one of these letters after all resolution steps indicate a failed
extraction and should be investigated.

## Known Data Quirks

- **Cross-job-code files (EJV/DRM):** EJV in 2026 contains 4 source files named `DRM 11.13.24...` and `DRM 12.3.24...` — historical files for the same physical project originally filed under DRM and re-released under EJV.  These contribute 102 rows to EJV.  Flag for accounting if cross-job deduplication matters.
- **Revision proliferation:** Jobs with many dated revisions (e.g. DAU with 4,515 rows, DEE with 2,787 rows) have inflated row counts because every revision is included by design.  Use `Source File` to group or deduplicate by release date in post-processing.
- **Summary-only jobs:** A small number of 2023 jobs have only Summary PDFs (no XML or ByStructure) — they predate the XML export workflow.

## Output File Status (last full run: 2026-05-15)

| File | Rows | Jobs | Invalid job codes | Notes |
|---|---|---|---|---|
| `bom_manhole_map.csv` | 18,365 | 151 | None | (2026) |
| `bom_manhole_map_2025.csv` | 81,392 | 214 | None | Re-run needed — ~171 folders missed due to 401s before proactive refresh fix |
| `bom_manhole_map_2024.csv` | 64,168 | 213 | None | Re-run needed to verify completeness |
| `bom_manhole_map_2023.csv` | 64,572 | 253 | None | |
| `skipped_structures.csv` | — | — | — | Overwritten on each run; lists structure names that matched no pattern |
| `errors.csv` | — | — | — | Overwritten on each run; NCF errors, BOM parse errors, folder-level errors |
| `unclassified_files.csv` | — | — | — | Overwritten on each run; files not matched by document_types.csv (noise extensions excluded) |

## Part Classification

`classify_part(category, part_number, description)` in `extract_all.py` determines `Part Type` and `Part Subtype`.

- **Resale** (early-exit before precast map): BOM category "Precast" AND product number starts with `PR` (Pro-Ring adjustment rings) or `HDPE` (Ladtech HDPE rings). These are purchased resale items, not manufactured precast. Subtype from `_PR_HDPE_SUBTYPE_MAP`: Pro-Ring, Ladtech, Grade, Finish, Flat. ~36,000 rows affected across all years.
- **Precast types** (Base, Section, Cone, Lid, etc.): BOM category "Precast", matched by product number prefix via `_PRECAST_TYPE_MAP`.
- **Resale**: all non-Precast BOM categories (Joint Seal, Frame & Ring, Hardware, Connectors, Miscellaneous).

## Document Type Detection

Document type is determined by `data/document_types.csv` — a lookup table of keyword + extension → label.
No detection logic is hardcoded in scripts; adding a new type requires only a new row in the CSV.

| keyword | extension | source_file_name |
|---|---|---|
| bom by structure | .xml | BOM by Structure XML |
| bom by structure | .pdf | BOM by Structure PDF |
| bom summary | .pdf | BOM Summary PDF |
| shop drawing | .pdf | Shop Drawing PDF |
| quotation | .pdf | Quotation PDF |

Priority within a job folder: XML > By Structure PDF > Summary PDF > Shop Drawing PDF.
All scripts load this table at startup via `_load_document_types()` / `classify_document()`.

Structure names extracted from Shop Drawing PDFs use the same `_STRUCTURE_PATTERNS` list as BOM parsers (20 patterns).
Any candidate name that fails all patterns is logged: `[SKIP structure] "name" — no pattern matched`.

Files not matched by `classify_document()` are logged to `output/unclassified_files.csv` (noise extensions like .docx, .msg, .jpg excluded).

## GitHub

Repository: `https://github.com/armorock/bom-structure-detail` (private)

## PDF Crash Isolation

PDF parsing uses subprocess wrappers (`parse_bom_pdf_safe`, `parse_bom_by_structure_pdf_safe`).
XML BOMs parse in-process (no crash risk).

## Coding Conventions

- No explanatory comments
- No premature abstractions
- Validate only at system boundaries
- All outputs to `output/` — use `os.makedirs("output", exist_ok=True)` before writing

## Future Options

**Multi-generation part name decoding**
The `/item-name-detail-interpreter` skill contains complete parsing rules for all 4 Armorock
naming generations (Gen1 2014–2017, Gen2 2017–2023, Gen3 2019–2023, Gen4 current).
The current `build_part_name()` and `classify_part()` in the extract scripts are Gen4-only —
older part numbers (e.g. `70424-B75`, `484S`, `S0472-0000`) fall back to the raw part number
as `Part Name` with no `Part Type`/`Part Subtype`.
To fix: port the generation detection + suffix mapping from the skill into `parse_bom_pdf.py`
as an updated `decode_part_number(pn)` function. Output columns stay the same — no schema change.
