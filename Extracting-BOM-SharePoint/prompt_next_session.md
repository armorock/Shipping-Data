# Session Prompt — BOM Extraction Continuation

## Project

Python script (`extract_bom_test.py`) pulls Armorock BOM files from SharePoint via
Microsoft Graph API, parses part data, and writes `output/bom_manhole_map.csv`.
Each row is one line item in one structure on one job. No files saved to disk.

SharePoint: `armorockllc.sharepoint.com/sites/JobData2026`, drive "Job Data 2026"

## File Roles

| File | Role |
|---|---|
| `graph_client.py` | Auth token, token caching, `graph_get`, `graph_get_all` |
| `sharepoint_client.py` | `get_site`, `get_drive`, `list_children`, `iter_files`, `download_file` |
| `parse_bom_pdf.py` | `parse_bom_pdf` (BOM Summary) + `parse_bom_by_structure_pdf` (PDF by Structure) + `parse_bom_by_structure_xml` (XML by Structure) |
| `extract_bom_test.py` | Main script → `output/bom_manhole_map.csv` |
| `extract_bom.py` | **Obsolete** — do not use |

## Output Columns (`output/bom_manhole_map.csv`)

`Year Release`, `BOM Release Date`, `Date extracted`, `Job Code`, `Project Name`,
`Structure Name`, `Job Location`, `Contractor`, `Agency`, `Engineer`,
`Part Name`, `Product Number`, `Quantity`, `Weight`, `Production Part`,
`Part Type`, `Part Subtype`

## What Was Done This Session

1. **XML BOM support** — added `parse_bom_by_structure_xml()` to `parse_bom_pdf.py`
   for SpreadsheetML files (`BOM by Structure (Excel XML).xml`)
2. **BOM priority order** — script now picks XML first, then PDF by-structure, then
   PDF summary as last resort; deduplicates structures across multiple files per job
3. **Antiquated Files excluded** — `sharepoint_client.py` skips `Antiquated Files/`
   subfolders so superseded BOMs are ignored
4. **Per-folder error isolation** — any folder/file error is logged and skipped;
   the run continues to completion regardless
5. **Token caching** — `graph_client.py` saves the refresh token to
   `~/.claude/msgraph_token.json`; subsequent runs skip the browser sign-in entirely
6. **Verbose folder logging** — skipped folders (regex no-match) now print
   `[no-match] <folder name>` so you can see why folders are being skipped
7. **Errors to stdout** — error messages go to stdout (not stderr) so they're
   always visible and captured by redirects

## How to Run

```powershell
py extract_bom_test.py
```

Token is cached — no sign-in prompt. To capture full output to a log file:

```powershell
py -u extract_bom_test.py | Tee-Object output\run_log.txt
```

(`-u` = unbuffered so output appears in real time)

## Current Issue Being Debugged

The script appears to stop after processing EGP (the first matching job folder) with
no output after. Suspected causes:

1. **Terminal truncation** — 147 folders × potential `[no-match]` lines may scroll
   off the terminal; actual run may be completing but user only sees the tail
2. **Silent crash after token acquisition** — `get_site()` may be failing and the
   exception disappears in the pipeline

**Next step**: run `py extract_bom_test.py` plain (token is cached, no sign-in needed)
and read the full terminal output, or check `output/bom_manhole_map.csv` row count.

If most folders show `[no-match]`, the job folder regex `^([A-Z]{3})\s*-\s*(.+)$`
may need widening (e.g. to allow 2–4 letter codes).

## Known Gaps (Not Yet Fixed)

| Gap | Detail |
|---|---|
| `Agency` / `Engineer` always blank | Come from New Customer Form, not yet parsed |
| Gasket rows: blank Part Name | PDF uses descriptive strings as part numbers; alphanumeric PN parser skips them |
| Location missing for job EHE | BOM header has no Location field; no fallback yet |
