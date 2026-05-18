# CLAUDE.md — Extracting M Drive

## Project Purpose

Extract every BOM file from the local M:\ drive (years 2016–2024) and produce per-year CSVs.
Same output schema as `BOM Structure Detail` but reads from local files — no SharePoint auth required.

## M:\ Drive Structure

| Year | Release Folder | Notes |
|------|----------------|-------|
| 2016 | `_Release` | Mostly PDFs; some have "Selected-BOM" in filename |
| 2017 | `_Release` | Sparse — few/no BOMs in many jobs |
| 2018 | `_Release` | PDF-only for most jobs |
| 2019 | `Release` or `_Release` | Full XML + PDF |
| 2020–2021 | `_Release` | Full XML + PDF |
| 2022 | `_release` (lowercase) | **Damaged/incomplete — data loss incident; only 10 jobs recovered** |
| 2023 | `_release` (lowercase) | Full XML + PDF |
| 2024 | `5-Release`, `(1) Release`, etc. | Varies per job |

Job folder naming: `CODE-Name`, `CODE - Name`, or `CODE- Name` — code is the first token.
Job code constraint: exactly 3 letters, first letter A–E (e.g. `BRA`, `CEH`). Folders that don't yield a valid code are skipped.

Release folder detection: any subfolder whose name contains "release" (case-insensitive).

## File Roles

| File | Role |
|---|---|
| `extract_mdrive.py` | Main script — local M:\ traversal, produces per-year CSVs |
| `parse_bom_pdf.py` | Shared parser module — keep in sync with BOM Structure Detail |
| `union_to_excel.py` | Combines M Drive 2016–2022 + BOM Structure Detail 2023–2026 into one XLSM |
| `combine_to_excel.py` | M Drive only — combines all per-year CSVs into a single XLSX (legacy) |
| `run_2016.py` … `run_2024.py` | Single-year runners — open 9 terminals and run in parallel |
| `data/location_overrides.csv` | Manual location fixes by job code |
| `data/state_abbreviations.csv` | State name → abbreviation lookup |
| `data/document_types.csv` | Lookup table: filename keyword + extension → Source File Name label |

## Usage

```
python run_2016.py          # run one year (open 9 terminals for all years in parallel)
python union_to_excel.py    # build output/all_bom_union.xlsm after all years complete
python combine_to_excel.py  # M Drive only XLSX (legacy)
```

Run from the project directory (`C:\Users\JohnLeitzke\Code\extracting M drive\`).

## Output Schema (22 columns)

Same as BOM Structure Detail — the two datasets union cleanly.
`Zip Code` is always blank (no NCF files on M: drive).
`Job Location` / `Location Source` come from the BOM header or `data/location_overrides.csv`.
`Source File Name` is the document type label from `data/document_types.csv` (e.g. `BOM by Structure XML`, `BOM Summary PDF`, `Quotation PDF`).

### Union output

`output/all_bom_union.xlsm` — 439,793 rows, all years 2016–2026:

| Source                                  | Years     | Rows    |
|-----------------------------------------|-----------|---------|
| M Drive (`extract_mdrive.py`)           | 2016–2022 | 139,120 |
| BOM Structure Detail (`extract_all.py`) | 2023–2026 | 300,673 |

2023–2024 are sourced from BOM Structure Detail only to avoid duplicate jobs.

## Part Classification

`Production Part` comes directly from the BOM category label. `Part Type` and `Part Subtype` are derived in `classify_part()`.

### Part Type values

| Part Type | Source |
|-----------|--------|
| Section, Base, Cone, Lid, Grade Ring, Taper Lid, Flat Top, Tee, Reducer, Rehab Ring, Box Culvert | Precast — matched by product number prefix via `_PRECAST_TYPE_MAP` or description keyword |
| Special | Precast — product number starts with `SPECIAL`; custom non-standard shapes |
| Resale | All non-Precast categories (Connectors, Joint Seal, Frame & Ring, Miscellaneous) **plus** PR* (Pro-Ring adjustment rings) and HDPE-* (Ladtech HDPE rings) even though their BOM category is `Precast` — purchased resale items, not manufactured precast |

### Pro-Ring and Ladtech (PR* and HDPE-* product numbers)
Pro-Ring adjustment rings and Ladtech HDPE rings appear under `Production Part = Precast` in the BOM but are classified as `Part Type = Resale` because they are purchased items. Subtype from `_PR_HDPE_SUBTYPE_MAP`: Pro-Ring, Ladtech, Grade, Finish, Flat. Handled by an early-exit check in `classify_part()` before the precast type map is consulted.

## Known Quirks

- `*-BOM 8.31.16.pdf` (one 2016 file with date in name) doesn't match BOM patterns — silently skipped
- Summary PDFs in 2016 sometimes use `Selected-BOM Summary.pdf` naming — handled in `_parse_summary_structure`
- 2016 PDFs may return empty structure names if the PDF template was different — check output rows
- Summary PDF rows have empty `Structure Name` by design — the format lists all parts combined
- ~14 rows across all years have `Part Name = IN` and no product number — parser artifact from truncated PDF lines; accepted as noise

## PDF Data Quality Issues (fixed in `parse_bom_pdf.py`)

| Symptom | Cause | Fix |
|---|---|---|
| `Production Part = PPrreeccaasstt` | Font encoding artifact — PDF reader doubled each character | `_dedouble()` collapses consecutive duplicate-character pairs before storing the category |
| `Production Part = DEPTH PLUS` | Section heading in Summary PDFs grouping custom-depth precast parts; weight also bleeds into part name | `_CATEGORY_ALIASES` maps `depth plus → Precast`; weight stripping now runs on these rows |
| Project name in `Production Part` | Long job title matched by `_ANY_CATEGORY_HEADER` as a category | Regex limited to ≤3 words; real categories are all ≤3 words (`Frame & Ring`, `Joint Seal`, etc.) |

## Output File Behavior

`skipped_structures.csv` is overwritten (fresh header) at the start of each `run_year()` call — every run produces a clean file, no accumulation of stale entries from prior runs.

## GitHub

Repository: `https://github.com/armorock/bom-mdrive` (private)
