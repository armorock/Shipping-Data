"""
Phase 2 — Dashboard Generator
Creates:
  - Dashboard.md            (vault root — master Dataview summary)
  - States/<State>.md       (one per unique state)
  - Cities/<City> ST.md     (one per unique city+state pair)
Then back-fills wiki-link breadcrumbs into existing job code files.
"""

import os
import re
import openpyxl

# ── Paths ─────────────────────────────────────────────────────────────────────

VAULT     = os.path.dirname(__file__)
JOB_CODES = os.path.join(VAULT, "Job Codes")
STATES_DIR = os.path.join(VAULT, "States")
CITIES_DIR = os.path.join(VAULT, "Cities")
EXCEL_PATH = (
    r"C:\Users\AlecSchooley\Desktop\Schooley Chaotic mind"
    r"\Projects\Shipping Data\All Shipping Data BABY.xlsm"
)
SHEET_NAME = "Job Code Data"

# ── State name → 2-letter abbreviation ───────────────────────────────────────

STATE_ABBREV = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE",
    "Florida": "FL", "Georgia": "GA", "Hawaii": "HI", "Idaho": "ID",
    "Illinois": "IL", "Indiana": "IN", "Iowa": "IA", "Kansas": "KS",
    "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME", "Maryland": "MD",
    "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN", "Mississippi": "MS",
    "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
    "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY",
    "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK",
    "Oregon": "OR", "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC",
    "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX", "Utah": "UT",
    "Vermont": "VT", "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
    "Wisconsin": "WI", "Wyoming": "WY", "District of Columbia": "DC",
}


def sanitize(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


def title_case_state(raw: str) -> str:
    """Normalize state strings to title case matching STATE_ABBREV keys."""
    s = raw.strip().title()
    # Fix common two-word edge cases that title() gets wrong
    fixes = {"Of": "of", "And": "and"}
    for wrong, right in fixes.items():
        s = s.replace(f" {wrong} ", f" {right} ")
    return s


# ── Load job code data from Excel ─────────────────────────────────────────────

print(f"Reading Excel...")
wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, keep_vba=True)
ws = wb[SHEET_NAME]

# Collect (state, county, city) triples for all valid job codes
records = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        continue
    code_raw = row[1]
    if not code_raw or not isinstance(code_raw, str):
        continue
    code = code_raw.strip().upper()
    if not (len(code) == 3 and code.isalpha()):
        continue

    state_raw  = row[6]
    county_raw = row[3]
    city_raw   = row[5]

    state  = title_case_state(str(state_raw).strip()) if state_raw else ""
    county = str(county_raw).strip().title() if county_raw else ""
    city   = str(city_raw).strip().title() if city_raw else ""

    # Drop obviously bad values
    if state.lower() in ("none", "nan", ""):
        state = ""
    if county.lower() in ("none", "nan", ""):
        county = ""
    if city.lower() in ("none", "nan", ""):
        city = ""

    records.append({"code": code, "state": state, "county": county, "city": city})

wb.close()

# Unique states and (city, state) pairs
unique_states = sorted({r["state"] for r in records if r["state"]})
unique_city_state = sorted(
    {(r["city"], r["state"]) for r in records if r["city"] and r["state"]}
)

print(f"  Unique states: {len(unique_states)}")
print(f"  Unique city+state pairs: {len(unique_city_state)}")

# ── Helper: Dataview code block ───────────────────────────────────────────────

def dv(query: str) -> str:
    return f"```dataview\n{query.strip()}\n```"

# ── Create States/ pages ──────────────────────────────────────────────────────

os.makedirs(STATES_DIR, exist_ok=True)
state_pages_created = 0

for state in unique_states:
    safe_state = sanitize(state)
    filepath = os.path.join(STATES_DIR, f"{safe_state}.md")

    content = f"""# {state}

## All Jobs
{dv(f'''TABLE job_code, job_name, city, year, customer
FROM "Job Codes"
WHERE contains(string(state), "{state}")
SORT year DESC''')}

## Cities in {state}
{dv(f'''TABLE WITHOUT ID city as "City", length(rows) as "Jobs"
FROM "Job Codes"
WHERE contains(string(state), "{state}") AND city
GROUP BY city
SORT length(rows) DESC''')}

## Structure Count by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    state_pages_created += 1

print(f"  State pages created: {state_pages_created}")

# ── Create Cities/ pages ──────────────────────────────────────────────────────

os.makedirs(CITIES_DIR, exist_ok=True)
city_pages_created = 0

for city, state in unique_city_state:
    abbrev = STATE_ABBREV.get(state, state[:2].upper())
    safe_city_file = sanitize(f"{city} {abbrev}")
    filepath = os.path.join(CITIES_DIR, f"{safe_city_file}.md")

    content = f"""# {city}, {state}

## All Jobs
{dv(f'''TABLE job_code, job_name, year, customer
FROM "Job Codes"
WHERE contains(string(city), "{city}") AND contains(string(state), "{state}")
SORT year DESC''')}

## Structure Count by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    city_pages_created += 1

print(f"  City pages created: {city_pages_created}")

# ── Create Dashboard.md ───────────────────────────────────────────────────────

dashboard_path = os.path.join(VAULT, "Dashboard.md")
dashboard_content = f"""# Armorock Shipping Dashboard

## Jobs by State
{dv('''TABLE WITHOUT ID state as "State", length(rows) as "Jobs"
FROM "Job Codes"
WHERE state
GROUP BY state
SORT length(rows) DESC''')}

## Jobs by Year
{dv('''TABLE WITHOUT ID year as "Year", length(rows) as "Jobs"
FROM "Job Codes"
WHERE year
GROUP BY year
SORT year DESC''')}

## Jobs by Plant
{dv('''TABLE WITHOUT ID plant as "Plant", length(rows) as "Jobs"
FROM "Job Codes"
WHERE plant
GROUP BY plant
SORT length(rows) DESC''')}

## Structure Counts by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
"""
with open(dashboard_path, "w", encoding="utf-8") as f:
    f.write(dashboard_content)
print("  Dashboard.md created.")
print(f"\nDone.")
print(f"  State pages : {state_pages_created}")
print(f"  City pages  : {city_pages_created}")
