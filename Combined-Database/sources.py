import os

REPO = r"C:\Users\JohnLeitzke\Code\Shipping-Data"
SCHOOLEY = os.path.join(REPO, "Schooleys Shit")
MASTER_CSV = os.path.join(SCHOOLEY, "MASTER CSV FILES")

BOM_UNION_XLSX = os.path.join(REPO, "Extracting-M-Drive", "output", "all_bom_union.xlsx")
BOM_UNION_SHEET = "BOM 2016-2026"
BOM_RESALE_SHEET = "Resale"

DISPATCH_CSV = os.path.join(MASTER_CSV, "Dispatch_Board_Master_2019-2025.csv")

BABY_XLSM = os.path.join(MASTER_CSV, "All Shipping Data BABY.xlsm")
BABY_SHEET = "Master List"

JOBCODE_DB_JSON = os.path.join(SCHOOLEY, "Jacks_Data_Improvement_Plans", "output", "jobcode_db.json")
JOB_CODE_REGISTRY_XLSX = os.path.join(SCHOOLEY, "Job_Code_Registry.xlsx")
REPAIR_LOG_CSV = os.path.join(SCHOOLEY, "jobcode_repair_log.csv")

GEONAMES_TXT = os.path.join(REPO, "Shipping-Map", "data", "US.txt")

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
OUTPUT_DIR = os.path.join(HERE, "output")
REVIEW_DIR = os.path.join(OUTPUT_DIR, "review")
SNAPSHOT_DIR = os.path.join(OUTPUT_DIR, "snapshots")
RESOLUTIONS_CSV = os.path.join(DATA_DIR, "resolutions.csv")
OBSERVATIONS_DB = os.path.join(OUTPUT_DIR, "observations.db")

# Overlap counting rule for 2019-2025 where Dispatch and ERP both record shipments.
#   "erp"      - ERP only, all years (accounting count of record; strict).
#   "dispatch" - Dispatch for 2019-2025, ERP elsewhere (cleanest grain but Dispatch is incomplete).
#   "union"    - ERP all years + pieces from jobs that appear ONLY in Dispatch (most complete,
#                no double-count since dispatch-only jobs are absent from ERP).
# The reconciliation showed Dispatch covers only ~21.5k of ~66.5k overlap pieces and adds just
# 51 unique jobs, so "union" is the accurate default. Flip after reading the reconciliation.
SHIPPED_OVERLAP = "union"
DISPATCH_YEARS = set(range(2019, 2026))

ERP_LABELS = {"QuickBooks": "erp_qb", "Fishbowl": "erp_fb", "Netsuite": "erp_ns"}

PLANTS = {"BC": "Boulder City", "SS": "Sulphur Springs", "PC": "Plant City"}


def ensure_dirs():
    for d in (DATA_DIR, OUTPUT_DIR, REVIEW_DIR, SNAPSHOT_DIR):
        os.makedirs(d, exist_ok=True)
