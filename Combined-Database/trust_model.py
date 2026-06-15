import csv
import json
import os
import re
from collections import Counter, defaultdict

import sources

# jobcode_db field -> the source-of-truth field name we use elsewhere
DB_FIELDS = {
    "shipping_city": "city",
    "shipping_state": "state",
    "shipping_zip": "zip",
    "customer": "customer",
    "plant": "plant",
}

# Sources that can supply each field, in the order we expose them when no signal exists.
DEFAULT_SOURCE_ORDER = ["bom_union", "dispatch", "erp_ns", "erp_qb", "erp_fb",
                        "jobcode_db", "registry", "geonames", "repair_log"]

# Map the source labels used inside jobcode_db conflict strings to our source ids.
DB_SOURCE_ALIASES = {
    "BOM": "bom_union", "DISPATCH": "dispatch", "SHIPPING": "erp_ns",
    "NETSUITE": "erp_ns", "MARKDOWN": "markdown", "NOTION": "notion",
    "REGISTRY": "registry",
}

CONFIDENCE_RANK = {"High": 3, "Manual": 3, "Medium": 2, "Low": 1}


def _alias(name):
    return DB_SOURCE_ALIASES.get(name.strip().upper(), name.strip().lower())


def _winner_from_source(field_source):
    if not field_source:
        return None
    return _alias(str(field_source).split("+")[0])


def _conflict_sources(conflict_str):
    # e.g. "BOM='NC'; Shipping='NO'; Markdown='NC'"
    return [_alias(m) for m in re.findall(r"([A-Za-z]+)\s*=", conflict_str or "")]


def learn():
    """Return {field: {source: {'wins':n,'losses':n,'trust':float}}} plus a repair summary."""
    recs = json.load(open(sources.JOBCODE_DB_JSON, encoding="utf-8"))
    wins = defaultdict(Counter)
    losses = defaultdict(Counter)
    for r in recs:
        for db_field, field in DB_FIELDS.items():
            res = r.get(db_field + "_resolution")
            if res != "CONFLICT":
                continue
            winner = _winner_from_source(r.get(db_field + "_source"))
            contenders = _conflict_sources(r.get(db_field + "_conflict"))
            if not winner or not contenders:
                continue
            wins[field][winner] += 1
            for c in contenders:
                if c != winner:
                    losses[field][c] += 1

    table = {}
    for field in DB_FIELDS.values():
        srcs = set(wins[field]) | set(losses[field])
        table[field] = {}
        for s in srcs:
            w, l = wins[field][s], losses[field][s]
            table[field][s] = {"wins": w, "losses": l, "trust": (w + 1) / (w + l + 2)}

    # Repair-log confidence calibration (for inferred fills).
    methods = Counter()
    confs = Counter()
    with open(sources.REPAIR_LOG_CSV, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("repair_method"):
                methods[row["repair_method"]] += 1
            if row.get("repair_confidence"):
                confs[row["repair_confidence"]] += 1
    return table, {"methods": dict(methods.most_common()), "confidence": dict(confs.most_common())}


_TABLE = None


def _table():
    global _TABLE
    if _TABLE is None:
        _TABLE, _ = learn()
    return _TABLE


def source_rank(field):
    """Sources for a field, best trust first; unseen sources fall back to DEFAULT_SOURCE_ORDER."""
    t = _table().get(field, {})
    seen = sorted(t, key=lambda s: t[s]["trust"], reverse=True)
    tail = [s for s in DEFAULT_SOURCE_ORDER if s not in seen]
    return seen + tail


def score(field, source, resolution=None):
    """Trust score in [0,1] for a (field, source) pair, nudged by resolution grade."""
    t = _table().get(field, {})
    base = t.get(source, {}).get("trust")
    if base is None:
        order = source_rank(field)
        base = 0.5 - 0.03 * order.index(source) if source in order else 0.3
    grade = {"MATCH": 1.0, "SINGLE_SOURCE": 0.85, "CONSENSUS_OVERRIDE": 0.9,
             "CONFLICT": 0.7, "NO_DATA": 0.5}.get((resolution or "").upper(), 1.0)
    return round(min(1.0, base) * grade, 4)


def write_report():
    sources.ensure_dirs()
    table, repair = learn()
    path = os.path.join(sources.OUTPUT_DIR, "trust_report.md")
    lines = ["# Source Trust Report",
             "",
             "Derived from `jobcode_db.json` conflict resolutions (which source supplied the",
             "winning value when sources disagreed) and `jobcode_repair_log.csv` confidence patterns.",
             ""]
    for field in DB_FIELDS.values():
        lines.append(f"## {field}")
        ranked = sorted(table[field].items(), key=lambda kv: kv[1]["trust"], reverse=True)
        if not ranked:
            lines.append("_no conflicts observed_\n")
            continue
        lines.append("| rank | source | trust | wins | losses |")
        lines.append("|---|---|---|---|---|")
        for i, (s, st) in enumerate(ranked, 1):
            lines.append(f"| {i} | {s} | {st['trust']:.3f} | {st['wins']} | {st['losses']} |")
        lines.append("")
    lines.append("## Repair-log confidence calibration")
    lines.append(f"- methods: {repair['methods']}")
    lines.append(f"- confidence tiers: {repair['confidence']}")
    lines.append("")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path


if __name__ == "__main__":
    p = write_report()
    print("wrote", p)
    for field in DB_FIELDS.values():
        print(field, "->", source_rank(field)[:4])
