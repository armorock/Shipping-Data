import csv
import os
from datetime import datetime

import sources

FIELDS = ["entity_key", "field", "chosen_value", "chosen_by", "confirmed_at"]


def load():
    """Return {(entity_key, field): chosen_value} of confirmed human decisions."""
    out = {}
    if not os.path.exists(sources.RESOLUTIONS_CSV):
        return out
    with open(sources.RESOLUTIONS_CSV, encoding="utf-8", newline="") as f:
        for r in csv.DictReader(f):
            ek, fld, val = r.get("entity_key"), r.get("field"), r.get("chosen_value")
            if ek and fld and val not in (None, ""):
                out[(ek.strip(), fld.strip())] = val.strip()
    return out


def upsert(records):
    """records: iterable of (entity_key, field, chosen_value, chosen_by). Merges into resolutions.csv."""
    sources.ensure_dirs()
    existing = {}
    if os.path.exists(sources.RESOLUTIONS_CSV):
        with open(sources.RESOLUTIONS_CSV, encoding="utf-8", newline="") as f:
            for r in csv.DictReader(f):
                existing[(r["entity_key"], r["field"])] = r
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    added = 0
    for ek, fld, val, who in records:
        if val in (None, ""):
            continue
        existing[(ek, fld)] = {"entity_key": ek, "field": fld, "chosen_value": val,
                               "chosen_by": who, "confirmed_at": stamp}
        added += 1
    with open(sources.RESOLUTIONS_CSV, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for row in existing.values():
            w.writerow(row)
    return added
