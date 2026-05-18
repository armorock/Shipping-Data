# Session Prompt — BOM Parser Accuracy Testing

## Project

Python script (`extract_bom_test.py`) pulls Armorock BOM PDFs from SharePoint via
Microsoft Graph API, parses part data with `parse_bom_pdf.py`, and writes
`output/bom_manhole_map.csv`. Each row is one line item in one structure on one job.
No files are saved to disk — PDFs are downloaded into memory, parsed, then discarded.

## Current output (last run: 2026-05-06)

- **2,608 rows** across **73 jobs** and **273 structures** (2026 drive only)
- Output columns: Year Release, BOM Release Date, Date extracted, Job Code,
  Project Name, Structure Name, Job Location, Contractor, Agency, Engineer,
  Part Name, Product Number, Quantity, Weight, Production Part, Part Type, Part Subtype

## Known uncertainty — 765 rows (29.3%) have at least one bad field

| Issue | Rows | % |
|---|---|---|
| Blank Part Name (parser missed part number in PDF) | 340 | 13.0% |
| Undecoded precast PN — PR\* and HDPE\* prefixes fall through to raw PN | 308 | 11.8% |
| Blank Part Type (same 308 rows as above) | 308 | 11.8% |
| Precast missing weight | 290 | 11.1% |
| Blank Job Location (job EHE only) | 26 | 1.0% |

The 1,843 fully-clean rows (70.7%) have all fields populated and decoded correctly.

## Known parser gaps (not yet implemented)

1. **Multi-PDF jobs**: script takes the single most recent BOM file per job folder.
   If a job has structures spread across multiple BOM PDFs, only the newest PDF's
   structures are captured. 273 structures is a floor, not a ceiling.
2. **Location fallback**: when BOM header has no Location field, no fallback yet.
   EHE is the only current gap.
3. **Agency / Engineer**: always blank — these come from the New Customer Form,
   which is not yet parsed.

## Goal for this session

Validate parser accuracy against an external source of truth.

**Opening question for Claude:**
I want to test the accuracy of this parser's output. I'm considering exporting data
from NetSuite as a comparison source — job codes, structure counts, part numbers,
and quantities are all trackable there. Is a NetSuite export a good validation
strategy, and what fields should I export to get the most coverage against the
CSV columns above? Are there gaps NetSuite won't cover that need a different source?
