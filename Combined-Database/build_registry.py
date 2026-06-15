import csv
import json
import os
import re
import sqlite3

import sources
import resolutions
import trust_model
from common import norm_job_code, is_blank

RESOLVED_FIELDS = ["city", "state", "zip", "customer", "plant"]
VALID_CODE = re.compile(r"[A-Z]{2,4}")
# Top trusted source per field; a conflict the top source backs is auto-resolved (no review).
TOP_SOURCE = {f: trust_model.source_rank(f)[0] for f in RESOLVED_FIELDS}
OUT_FIELDS = ["job_code", "project_name", "customer", "contractor", "city", "state",
              "zip", "county", "plant", "year",
              "city_status", "state_status", "zip_status", "customer_status", "plant_status",
              "sources", "needs_review"]


def _winners(con):
    out = {}
    for ek, field, status, winner, winner_sources in con.execute(
            "SELECT entity_key, field, status, winner, winner_sources FROM entity_field_resolution"):
        out.setdefault(ek, {})[field] = (winner, status, winner_sources or "")
    return out


def _jobcode_db_extras():
    """county/contractor/project_name/year per job from jobcode_db (low-conflict supplements)."""
    recs = json.load(open(sources.JOBCODE_DB_JSON, encoding="utf-8"))
    out = {}
    for r in recs:
        jc = norm_job_code(r.get("job_code"))
        if not jc:
            continue
        out[jc] = {
            "county": None if is_blank(r.get("shipping_county")) else str(r["shipping_county"]).strip(),
            "contractor": None if is_blank(r.get("contractor")) else str(r["contractor"]).strip(),
            "project_name": None if is_blank(r.get("project_name")) else str(r["project_name"]).strip(),
            "year": r.get("year_released"),
        }
    return out


def build():
    sources.ensure_dirs()
    con = sqlite3.connect(sources.OBSERVATIONS_DB)
    winners = _winners(con)
    src_map = {}
    for ek, srcs in con.execute(
            "SELECT entity_key, GROUP_CONCAT(DISTINCT source) FROM observations GROUP BY entity_key"):
        src_map[ek] = srcs
    con.close()

    confirmed = resolutions.load()
    extras = _jobcode_db_extras()
    # Keep real job codes only: 2-4 letter codes, or anything jobcode_db knows about.
    job_codes = sorted(jc for jc in (set(winners) | set(extras))
                       if VALID_CODE.fullmatch(jc) or jc in extras)

    rows = []
    review_count = 0
    for jc in job_codes:
        w = winners.get(jc, {})
        ex = extras.get(jc, {})
        rec = {"job_code": jc,
               "project_name": ex.get("project_name"),
               "contractor": ex.get("contractor"),
               "county": ex.get("county"),
               "year": (w.get("year", (None,))[0] or ex.get("year")),
               "sources": src_map.get(jc, "")}
        needs = False
        for field in RESOLVED_FIELDS:
            winner, status, winner_sources = w.get(field, (None, "GAP", ""))
            value = confirmed.get((jc, field), winner)
            if (jc, field) in confirmed:
                status = "CONFIRMED"
            elif status == "CONFLICT" and TOP_SOURCE[field] in winner_sources.split(","):
                # Top-trust source backs the winner -> confidently resolved, not review-worthy.
                status = "RESOLVED"
            rec[field] = value
            rec[field + "_status"] = status
            if status in ("CONFLICT", "GAP") and field in ("city", "state"):
                needs = True
        rec["needs_review"] = "Y" if needs else ""
        if needs:
            review_count += 1
        rows.append(rec)

    path = os.path.join(sources.OUTPUT_DIR, "project_registry.csv")
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=OUT_FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k) for k in OUT_FIELDS})
    print(f"project_registry: {len(rows)} job codes | needs_review: {review_count}")
    print("wrote", path)
    return rows


if __name__ == "__main__":
    build()
