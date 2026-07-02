"""
Quarterly update entry point.

Prerequisites (manual, require network access):
  - Download NetSuite "AppxShipped" saved search as XLS -> Combined-Database/data/AppxShipped*.xls
  - Download NetSuite "OSLocationData" saved search as XLS -> Combined-Database/data/OSLocationData*.xls
  - BOM SharePoint extraction and M-drive extraction if source data changed (run those scripts separately)

Steps run here:
  [1] ingest_netsuite.py  — appends new NetSuite records to NSAW xlsx, flags geo anomalies
  [2] build_map_data.py   — geocodes, writes output/data.json + output/index.html

Usage:
  python run_quarterly_update.py
  python run_quarterly_update.py --skip-ingest   # skip step 1, just rebuild map
"""

import argparse
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent
SHIPPING_MAP = ROOT / "Shipping-Map"
XLSX_PATH = Path(r"C:\Users\JohnLeitzke\OneDrive - Armorock LLC\Documents\Desktop\NSAW All Shipping Data1.2.xlsx")
JOBCODE_DB = ROOT / "Schooleys Shit" / "Jacks_Data_Improvement_Plans" / "output" / "jobcode_db.json"


def check(path, label):
    if not path.exists():
        sys.exit(f"ERROR: {label} not found:\n  {path}")


def run(script, cwd):
    result = subprocess.run([sys.executable, str(script)], cwd=str(cwd))
    if result.returncode != 0:
        sys.exit(f"\nERROR: {script.name} failed (exit {result.returncode})")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-ingest", action="store_true", help="Skip NetSuite ingest, just rebuild map")
    args = parser.parse_args()

    check(XLSX_PATH, "NSAW xlsx")
    check(JOBCODE_DB, "jobcode_db.json")

    print("=" * 60)
    print("Armorock Shipping Map — Quarterly Update")
    print("=" * 60)
    print()
    print("NOTE: BOM SharePoint and M-drive extraction are prerequisites.")
    print("      Run those manually if source data has changed.")
    print()

    if not args.skip_ingest:
        print("[1/2] Ingesting new NetSuite records...")
        run(SHIPPING_MAP / "ingest_netsuite.py", SHIPPING_MAP)
        print()
    else:
        print("[1/2] Skipping NetSuite ingest (--skip-ingest)")
        print()

    print("[2/2] Building map data...")
    run(SHIPPING_MAP / "build_map_data.py", SHIPPING_MAP)
    print()

    print("=" * 60)
    print("Done. To deploy to GitHub Pages:")
    print()
    print("  cd Shipping-Map/output")
    print('  git add data.json index.html && git commit -m "Q refresh" && git push')
    print("=" * 60)


if __name__ == "__main__":
    main()
