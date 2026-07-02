import csv
import difflib
import json
import os
import random
import shutil
from collections import Counter, defaultdict
from datetime import date, datetime

import openpyxl

XLSX_PATH = r"C:\Users\JohnLeitzke\OneDrive - Armorock LLC\Documents\Desktop\NSAW All Shipping Data1.2.xlsx"
JOBCODE_DB_PATH = r"C:\Users\JohnLeitzke\Code\Shipping-Data\Schooleys Shit\Jacks_Data_Improvement_Plans\output\jobcode_db.json"
GEONAMES_PATH = os.path.join("data", "US.txt")

PLANTS = ["Boulder City", "Sulphur Springs", "Plant City"]

STATE_ABBREV = {
    "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR",
    "CALIFORNIA": "CA", "COLORADO": "CO", "CONNECTICUT": "CT", "DELAWARE": "DE",
    "DISTRICT OF COLUMBIA": "DC", "FLORIDA": "FL", "GEORGIA": "GA", "HAWAII": "HI",
    "IDAHO": "ID", "ILLINOIS": "IL", "INDIANA": "IN", "IOWA": "IA", "KANSAS": "KS",
    "KENTUCKY": "KY", "LOUISIANA": "LA", "MAINE": "ME", "MARYLAND": "MD",
    "MASSACHUSETTS": "MA", "MICHIGAN": "MI", "MINNESOTA": "MN", "MISSISSIPPI": "MS",
    "MISSOURI": "MO", "MONTANA": "MT", "NEBRASKA": "NE", "NEVADA": "NV",
    "NEW HAMPSHIRE": "NH", "NEW JERSEY": "NJ", "NEW MEXICO": "NM", "NEW YORK": "NY",
    "NORTH CAROLINA": "NC", "NORTH DAKOTA": "ND", "OHIO": "OH", "OKLAHOMA": "OK",
    "OREGON": "OR", "PENNSYLVANIA": "PA", "RHODE ISLAND": "RI",
    "SOUTH CAROLINA": "SC", "SOUTH DAKOTA": "SD", "TENNESSEE": "TN", "TEXAS": "TX",
    "UTAH": "UT", "VERMONT": "VT", "VIRGINIA": "VA", "WASHINGTON": "WA",
    "WEST VIRGINIA": "WV", "WISCONSIN": "WI", "WYOMING": "WY",
}

STATE_CENTROIDS = {
    "AL": (32.79, -86.83), "AK": (64.07, -152.28), "AZ": (34.27, -111.66),
    "AR": (34.89, -92.44), "CA": (37.18, -119.47), "CO": (39.00, -105.55),
    "CT": (41.62, -72.73), "DE": (38.99, -75.51), "DC": (38.91, -77.01),
    "FL": (28.63, -82.45), "GA": (32.64, -83.44), "HI": (20.29, -156.37),
    "ID": (44.35, -114.61), "IL": (40.04, -89.20), "IN": (39.89, -86.28),
    "IA": (42.07, -93.50), "KS": (38.49, -98.38), "KY": (37.53, -85.30),
    "LA": (31.07, -92.00), "ME": (45.37, -69.24), "MD": (39.06, -76.80),
    "MA": (42.26, -71.81), "MI": (44.35, -85.41), "MN": (46.28, -94.31),
    "MS": (32.74, -89.67), "MO": (38.35, -92.46), "MT": (47.05, -109.63),
    "NE": (41.54, -99.80), "NV": (39.33, -116.63), "NH": (43.68, -71.58),
    "NJ": (40.19, -74.67), "NM": (34.41, -106.11), "NY": (42.95, -75.53),
    "NC": (35.56, -79.39), "ND": (47.45, -100.47), "OH": (40.29, -82.79),
    "OK": (35.59, -97.49), "OR": (43.93, -120.56), "PA": (40.88, -77.80),
    "RI": (41.68, -71.56), "SC": (33.92, -80.90), "SD": (44.44, -100.23),
    "TN": (35.86, -86.35), "TX": (31.48, -99.33), "UT": (39.31, -111.67),
    "VT": (44.07, -72.67), "VA": (37.52, -78.85), "WA": (47.38, -120.45),
    "WV": (38.64, -80.62), "WI": (44.62, -89.99), "WY": (43.00, -107.55),
}

MHB_BASELINE = {
    2014: 255, 2015: 383, 2016: 767, 2017: 1176, 2018: 1353, 2019: 1431,
    2020: 1427, 2021: 1764, 2022: 1387, 2023: 1603, 2024: 2145, 2025: 2123,
    2026: 1025,
}

CITY_PREFIX_EXPANSIONS = [("ST ", "SAINT "), ("FT ", "FORT "), ("MT ", "MOUNT ")]


def norm_city(v):
    if not isinstance(v, str) or not v.strip():
        return None
    s = " ".join(v.upper().replace(".", " ").replace(",", " ").split())
    return s or None


def norm_state(v):
    if not isinstance(v, str):
        return None
    s = v.strip().upper()
    if s in STATE_CENTROIDS:
        return s
    return STATE_ABBREV.get(s)


def norm_zip(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        s = str(int(v)).zfill(5)
    else:
        s = str(v).strip().split("-")[0]
    return s if len(s) == 5 and s.isdigit() else None


def clean_year(v):
    if isinstance(v, (datetime, date)):
        return v.year
    return None


def clean_month(v):
    if isinstance(v, (datetime, date)):
        return v.month
    return None


def norm_job_code(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s and s.lower() != "none" else None


def load_geonames():
    zip_ll = {}
    zip_state = {}
    city_points = defaultdict(list)
    city_to_states = defaultdict(set)
    with open(GEONAMES_PATH, encoding="utf-8") as f:
        for row in csv.reader(f, delimiter="\t"):
            zip5, place, state = row[1], row[2], row[4]
            try:
                lat, lng = float(row[9]), float(row[10])
            except ValueError:
                continue
            if state not in STATE_CENTROIDS:
                continue
            zip_ll[zip5] = (lat, lng)
            zip_state[zip5] = state
            city_points[(place.upper(), state)].append((lat, lng))
            city_to_states[place.upper()].add(state)
    city_ll = {
        k: (sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts))
        for k, pts in city_points.items()
    }
    city_names_by_state = defaultdict(list)
    for city, state in city_ll:
        city_names_by_state[state].append(city)
    return zip_ll, zip_state, city_ll, city_names_by_state, city_to_states


def load_jobdb():
    with open(JOBCODE_DB_PATH, encoding="utf-8") as f:
        records = json.load(f)
    return {r["job_code"].strip(): r for r in records if r.get("job_code")}


def read_rows():
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    stats = Counter()
    header = None
    for raw in ws.iter_rows(values_only=True):
        if header is None:
            header = {str(h).strip(): i for i, h in enumerate(raw) if h is not None}
            continue
        qty = raw[header["Quantity"]]
        if qty == 0:
            stats["qty_zero_excluded"] += 1
            continue
        raw_date = raw[header["Date Shipped"]]
        year = clean_year(raw_date)
        month = clean_month(raw_date)
        if year is None:
            stats["bad_date_excluded"] += 1
            continue
        plant = str(raw[header["Plant"]]).strip() if raw[header["Plant"]] else None
        if plant not in PLANTS:
            stats["bad_plant_excluded"] += 1
            continue
        diameter = raw[header["diameter"]]
        rows.append({
            "plant_idx": PLANTS.index(plant),
            "customer": (raw[header["Invoiced Custumer"]] or "").strip() if isinstance(raw[header["Invoiced Custumer"]], str) else None,
            "job_code": norm_job_code(raw[header["Job Code"]]),
            "year": year,
            "month": month,
            "structure_name": raw[header["Structure Name"]].strip() if isinstance(raw[header["Structure Name"]], str) and raw[header["Structure Name"]].strip().lower() not in ("", "none") else None,
            "part_number": str(raw[header["Part Number"]]).strip() if raw[header["Part Number"]] else None,
            "part_type": raw[header["part_type"]],
            "diameter": str(diameter).strip() if diameter and str(diameter).strip().lower() != "none" else None,
            "city": norm_city(raw[header["Shipping City"]]),
            "state": norm_state(raw[header["Shippings State"]]),
            "zip": norm_zip(raw[header["ZipCode"]]),
            "part_name": str(raw[header["Current Part Name"]]).strip() if raw[header["Current Part Name"]] and str(raw[header["Current Part Name"]]).strip() not in ("None", "#MISSING") else None,
        })
    stats["rows_kept"] = len(rows)
    return rows, stats


def enrich(rows, jobdb):
    filled = Counter()
    for r in rows:
        db = jobdb.get(r["job_code"]) if r["job_code"] else None
        if db:
            if not r["state"] and norm_state(db.get("shipping_state")):
                r["state"] = norm_state(db["shipping_state"])
                filled["state"] += 1
            if not r["city"] and norm_city(db.get("shipping_city")):
                r["city"] = norm_city(db["shipping_city"])
                filled["city"] += 1
            if not r["zip"] and norm_zip(db.get("shipping_zip")):
                r["zip"] = norm_zip(db["shipping_zip"])
                filled["zip"] += 1
            r["project"] = db.get("project_name")
            r["contractor"] = db.get("contractor")
            if not r["customer"]:
                r["customer"] = db.get("customer")
        else:
            r["project"] = None
            r["contractor"] = None
    return filled


class Geocoder:
    def __init__(self, zip_ll, zip_state, city_ll, city_names_by_state, city_to_states):
        self.zip_ll = zip_ll
        self.zip_state = zip_state
        self.city_ll = city_ll
        self.city_names_by_state = city_names_by_state
        self.city_to_states = city_to_states
        self.fuzzy_cache = {}
        self.fuzzy_log = []

    def _is_suspect(self, city, state):
        if not city or not state:
            return False
        valid = self.city_to_states.get(city.upper(), set())
        return bool(valid) and state not in valid

    def lookup(self, city, state, zip5):
        if zip5 and zip5 in self.zip_ll:
            return self.zip_ll[zip5] + (0,)
        if state and city:
            variants = [city]
            for pre, full in CITY_PREFIX_EXPANSIONS:
                if city.startswith(pre):
                    variants.append(full + city[len(pre):])
                if city.startswith(full):
                    variants.append(pre + city[len(full):])
            for v in variants:
                if (v, state) in self.city_ll:
                    return self.city_ll[(v, state)] + (1,)
            key = (city, state)
            if key not in self.fuzzy_cache:
                matches = difflib.get_close_matches(city, self.city_names_by_state.get(state, []), n=1, cutoff=0.85)
                self.fuzzy_cache[key] = matches[0] if matches else None
                if matches:
                    self.fuzzy_log.append(f"  fuzzy: {city}, {state} -> {matches[0]}")
            if self.fuzzy_cache[key]:
                return self.city_ll[(self.fuzzy_cache[key], state)] + (1,)
        if state:
            return STATE_CENTROIDS[state] + (2,)
        return None


def main():
    os.makedirs("output", exist_ok=True)
    zip_ll, zip_state, city_ll, city_names_by_state, city_to_states = load_geonames()
    geo = Geocoder(zip_ll, zip_state, city_ll, city_names_by_state, city_to_states)
    jobdb = load_jobdb()
    rows, stats = read_rows()
    filled = enrich(rows, jobdb)

    locs = []
    loc_index = {}
    prec_counts = Counter()
    anomalies = []
    for r in rows:
        result = geo.lookup(r["city"], r["state"], r["zip"])
        if result is None:
            r["loc_idx"] = None
            r["prec"] = None
            prec_counts["unmapped"] += 1
            continue
        lat, lng, prec = round(result[0], 4), round(result[1], 4), result[2]
        if prec != 3 and geo._is_suspect(r["city"], r["state"]):
            prec = 3
            anomalies.append({
                "job_code":     r.get("job_code") or "",
                "city":         r["city"] or "",
                "listed_state": r["state"] or "",
                "valid_states": "|".join(sorted(geo.city_to_states.get((r["city"] or "").upper(), set()))),
                "zip":          r["zip"] or "",
                "year":         str(r["year"]),
                "part_type":    r.get("part_type") or "",
            })
        key = (lat, lng)
        if key not in loc_index:
            loc_index[key] = len(locs)
            locs.append([lat, lng])
        r["loc_idx"] = loc_index[key]
        r["prec"] = prec
        prec_counts[prec] += 1

    customers = []
    cust_index = {}
    jobs = []
    job_index = {}
    job_rows = defaultdict(list)
    for r in rows:
        key = r["job_code"] or ("", r["customer"] or "", r["city"] or "", r["state"] or "")
        job_rows[key].append(r)
    for key, grp in job_rows.items():
        first = grp[0]
        cust = first["customer"]
        if cust not in cust_index:
            cust_index[cust] = len(customers)
            customers.append(cust)
        comp = Counter(r["part_name"] or r["part_number"] for r in grp if r["part_name"] or r["part_number"])
        years = [r["year"] for r in grp]
        jobs.append([
            first["job_code"] or "",
            first["project"],
            cust_index[cust],
            first["contractor"],
            first["city"].title() if first["city"] else None,
            first["state"],
            first["zip"],
            [[p, n] for p, n in comp.most_common(15)],
            [min(years), max(years)],
        ])
        job_index[key] = len(jobs) - 1
        for r in grp:
            r["job_idx"] = job_index[key]

    mhb_agg = defaultdict(lambda: {"qty": 0, "diams": Counter()})
    mhb_unmapped = 0
    for r in rows:
        if r["part_type"] != "MHB":
            continue
        if r["loc_idx"] is None:
            mhb_unmapped += 1
            continue
        key = (r["loc_idx"], r["year"], r["month"], r["plant_idx"], r["job_idx"], r["prec"])
        mhb_agg[key]["qty"] += 1
        if r["diameter"]:
            mhb_agg[key]["diams"][r["diameter"]] += 1
    mhb = [
        [loc, y, mo, p, v["qty"], j, prec, [[d, n] for d, n in sorted(v["diams"].items(), key=lambda x: -x[1])]]
        for (loc, y, mo, p, j, prec), v in mhb_agg.items()
    ]

    data = {
        "meta": {
            "generated": date.today().isoformat(),
            "currentMonth": date.today().month,
            "notMappedMhb": mhb_unmapped,
            "notMappedRows": prec_counts["unmapped"],
            "badDates": stats["bad_date_excluded"],
            "years": sorted(MHB_BASELINE),
            "plants": PLANTS,
        },
        "locs": locs,
        "customers": customers,
        "jobs": jobs,
        "mhb": mhb,
    }
    out_path = os.path.join("output", "data.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, separators=(",", ":"), ensure_ascii=False)

    if os.path.exists("index.html"):
        shutil.copy("index.html", os.path.join("output", "index.html"))
    with open(os.path.join("output", "robots.txt"), "w") as f:
        f.write("User-agent: *\nDisallow: /\n")

    anom_path = os.path.join("output", "geo_anomalies.csv")
    if anomalies:
        with open(anom_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["job_code", "city", "listed_state", "valid_states", "zip", "year", "part_type"])
            writer.writeheader()
            writer.writerows(anomalies)

    print("=== VERIFICATION ===")
    print(f"rows kept: {stats['rows_kept']} | qty=0 excluded: {stats['qty_zero_excluded']} | bad dates: {stats['bad_date_excluded']} | bad plant: {stats['bad_plant_excluded']}")
    print(f"enrichment fills from jobcode_db: {dict(filled)}")
    mhb_by_year = Counter()
    for loc, y, mo, p, qty, j, prec, diams in mhb:
        mhb_by_year[y] += qty
    ok = True
    for y in sorted(MHB_BASELINE):
        actual = mhb_by_year[y]
        mark = "OK" if actual + (0) <= MHB_BASELINE[y] else "??"
        print(f"  {y}: mapped MHB {actual} / baseline {MHB_BASELINE[y]}")
        if actual > MHB_BASELINE[y]:
            ok = False
    total_mapped = sum(mhb_by_year.values())
    print(f"MHB mapped: {total_mapped} + unmapped: {mhb_unmapped} = {total_mapped + mhb_unmapped} (baseline {sum(MHB_BASELINE.values())})")
    if total_mapped + mhb_unmapped != sum(MHB_BASELINE.values()):
        ok = False
    print(f"baseline check: {'PASS' if ok else 'FAIL'}")
    print(f"geocode precision: zip={prec_counts[0]} city={prec_counts[1]} state-centroid={prec_counts[2]} suspect={prec_counts[3]} unmapped={prec_counts['unmapped']}")
    if anomalies:
        print(f"WARN: {len(anomalies)} suspect city/state rows -> {anom_path}")
    print(f"fuzzy city matches ({len(geo.fuzzy_log)}):")
    for line in geo.fuzzy_log:
        print(line)
    print(f"locs: {len(locs)} | jobs: {len(jobs)} | mhb records: {len(mhb)}")
    print(f"data.json: {os.path.getsize(out_path):,} bytes")
    sample = random.sample([r for r in rows if r["loc_idx"] is not None], 5)
    print("spot-check geocodes:")
    for r in sample:
        lat, lng = locs[r["loc_idx"]]
        print(f"  {r['job_code']} | {r['city']}, {r['state']} {r['zip']} -> {lat},{lng} (prec {r['prec']})")


if __name__ == "__main__":
    main()
