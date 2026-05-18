# Extracting BOM Job Data from SharePoint

Pulls Armorock BOM Summary PDFs from the company SharePoint and extracts part data into CSVs for manhole location mapping by job code.

## Prerequisites

```
pip install requests pdfplumber
```

Microsoft Graph credentials must be stored in `~/.claude/msgraph_config.json`:
```json
{
  "tenant_id": "YOUR_TENANT_ID",
  "client_id": "YOUR_CLIENT_ID"
}
```
Authentication uses the device-code flow — you'll be prompted to visit a URL and sign in once per run.

## Files

| File | Purpose |
|---|---|
| `graph_client.py` | Microsoft Graph authentication (device-code flow) and HTTP helpers |
| `sharepoint_client.py` | SharePoint site/drive traversal and file download |
| `parse_bom_pdf.py` | Extracts header, line items, and opening schedule from a BOM Summary PDF |
| `extract_bom.py` | Full extraction → `bom_line_items.csv` + `bom_openings.csv` |
| `extract_bom_test.py` | Test extraction → `bom_manhole_map.csv` (matches Example Formatting.csv columns) |
| `Example Formatting.csv` | Reference column layout for the target output format |

## Quickstart — Test Extraction

```
python extract_bom_test.py
```

Outputs `bom_manhole_map.csv` with one row per Precast line item per BOM PDF.

## Quickstart — Full Extraction

```
python extract_bom.py
```

Outputs two CSVs:
- `bom_line_items.csv` — all line items (all categories) keyed by job code
- `bom_openings.csv` — pipe openings schedule keyed by job code

## Output Columns (`bom_manhole_map.csv`)

| Column | Source |
|---|---|
| Year Release | Date embedded in BOM filename (e.g. `4.20.26` → `2026`) |
| Job Code | 3-letter prefix from job folder name (e.g. `ELC`) |
| Project Name | `Job Name` field from PDF header |
| Structure Name | Segment of BOM filename between job name and `-BOM Summary.pdf` (e.g. `SSMH-001`, `Selected`) |
| Job Location | `Location` field from PDF header |
| Contractor | `Contractor` field from PDF header |
| Agency | Blank — not present in BOM PDFs |
| Engineer | Blank — not present in BOM PDFs |
| Part Description | Line item description text |
| Product Number | Line item part number |
| Quantity | Line item quantity |
| Weight | Line item weight (lbs), Precast items only |
| Part Type | Derived from description keywords (Cone, Base, Riser, Lid, etc.) |
| Part Subtype | Derived from description keywords (Eccentric, Flat, Concentric, etc.) |

## SharePoint Configuration

The 2026 site settings (confirmed):

```python
HOSTNAME   = "armorockllc.sharepoint.com"
SITE_PATH  = "/sites/JobData2026"
DRIVE_NAME = "Job Data 2026"
JOBS_ROOT  = "root"
```

Job folders sit at the drive root and follow the pattern `XXX - Job Name` (3-letter code, dash, name). BOM PDFs are inside subfolders (`5-Release/`, etc.) and are found recursively.
