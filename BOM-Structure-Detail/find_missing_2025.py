"""
Find 2025 job folders on SharePoint that are absent from all_bom_2025.csv.
These are likely the folders that failed with 401 FOLDER ERRORs during extraction.

Run: python find_missing_2025.py
"""
import csv
import re
from graph_client import acquire_token
from sharepoint_client import get_site, get_drive, list_children

HOSTNAME   = "armorockllc.sharepoint.com"
SITE_PATH  = "/sites/jobdata2025"
DRIVE_NAME = "Job Data 2025"
CSV_PATH   = "output/all_bom_2025.csv"

_JOB_FOLDER_RE = re.compile(r"^([A-Z]{2,4})\s*-\s*", re.IGNORECASE)

token = acquire_token()
site  = get_site(token, HOSTNAME, SITE_PATH)
drive = get_drive(token, site["id"], DRIVE_NAME)

print(f"Listing folders in {DRIVE_NAME}...")
children     = list_children(token, drive["id"], "root")
sp_folders   = [c["name"] for c in children if "folder" in c]

sp_codes = {}
for name in sp_folders:
    m = _JOB_FOLDER_RE.match(name.strip())
    code = m.group(1).upper() if m else name.strip().upper()
    sp_codes[code] = name

print(f"  {len(sp_folders)} folders on SharePoint")

csv_codes = set()
with open(CSV_PATH, newline="", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        code = row.get("Job Code", "").strip().upper()
        if code:
            csv_codes.add(code)

print(f"  {len(csv_codes)} job codes in {CSV_PATH}\n")

missing = {code: name for code, name in sp_codes.items() if code not in csv_codes}

if not missing:
    print("No missing folders — output looks complete.")
else:
    print(f"{len(missing)} folders missing from CSV (likely 401 victims):\n")
    for code, name in sorted(missing.items()):
        print(f"  {name}")
