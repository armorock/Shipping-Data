# Combined-Database — project notes

Standalone project that combines Armorock's upstream extraction sources into one auditable source of
truth. Replaces the deprecated `NSAW All Shipping Data1.2.xlsx` mashup. See `README.md` for usage.

## Architecture (module map)

- `sources.py` — every input path + config (`SHIPPED_OVERLAP`, dirs). Change paths here only.
- `common.py` — normalization + `structure_class()` (multi-generation part-number classifier).
- `trust_model.py` — learns field x source trust from `jobcode_db.json` conflict resolutions +
  `jobcode_repair_log.csv`. `score(field, source, resolution)` and `source_rank(field)` used elsewhere.
- `build_observation_db.py` — reads all sources -> `observations` table (one row per reported value,
  with `obs_id`, source, trust) + `entity_field_resolution` view (AGREE/CONFLICT/GAP + winner).
- `resolutions.py` — durable confirmed decisions in `data/resolutions.csv` (load/upsert).
- `build_registry.py` — Layer 1 (`project_registry.csv`). A CONFLICT whose winner is backed by the
  top-trust source is auto-resolved (`RESOLVED`); only city/state CONFLICT/GAP set `needs_review`.
- `build_pieces_ledger.py` — Layer 2 (`pieces_ledger.csv`). ERP dedup = max-count-per-system per
  (job, part, date). Overlap rule from `sources.SHIPPED_OVERLAP`. Emits reconciliation + quoted_not_shipped.
- `build_reports.py` — rehab/nonbase report, provenance, structure completeness roadmap, review lists.
  Review list columns: `job_code | field | status | current_winner | winner_sources | project_name | [field extras] | src_bom_union | src_dispatch | src_erp_qb | src_erp_fb | src_erp_ns | src_jobcode_db | src_registry | all_observed_values | Confirmed Value | Notes`.
  City extras: `raw_bom | zip_county | county_flag`. Zip extras: `geo_city | geo_state | confirmed_city | confirmed_state | zip_ok` (GeoNames from `Shipping-Map/data/US.txt`).
  Also produces: `04_county.xlsx` (county CONFLICT/GAP rows) and `07_location_matrix.xlsx` (all jobs, all sources, includes AGREE rows + `street_address` column for geocodable addresses).
- `apply_corrections.py` — reads filled review xlsx -> resolutions -> rebuilds registry/ledger/reports.
- `run_quarter.py` — orchestrator. `snapshot_and_diff.py` — per-quarter snapshot + change report.

## Decisions / gotchas

- **`county` is a separate field** from `city` — county-in-city values (e.g., "CLARK COUNTY") are rerouted to the county field via `is_county_value()`. Three sources have dedicated county columns now read: BABY (`Shipping County`), Registry (`County`), jobcode_db (`shipping_county`).
- **`project_name` field** captured from BOM (`Project Name` col), Dispatch (`job_name`), Registry (`BOM Project Name`), jobcode_db (`project_name`). Shows as context column in all review exports; excluded from 06_conflicts_all (variant spellings create noise).
- **`street_address` field** — BOM "Job Location" rows where the state parse fails (e.g., full delivery addresses) are stored here instead of discarded. Higher-confidence location data; available for future geocoding.
- **`norm_city()`** strips trailing state codes ("PHOENIX AZ" → "PHOENIX") and rejects bare state codes. **`norm_zip()`** rejects "00000". **`norm_county()`** strips trailing state codes and the word "COUNTY".
- **Classification priority is part-number FIRST**, then the human `part_type` word. Dispatch labels
  RMH (rehab) parts as "section", so word-first undercounts rehabs (1,427 vs the correct 3,304).
- **A piece must have a part number.** Blank-qty ERP rows with a part = 1 piece; rows with no part
  number are empty/junk and excluded.
- **ERP systems barely overlap** (only 253 of 45k (job,part,date) groups span >1 system), so cross-ERP
  duplication is small; the real overlap is Dispatch vs ERP (handled by `SHIPPED_OVERLAP`).
- **Dispatch is incomplete** for 2019–2025 (esp. 2019–2021); ERP is the count of record. Default
  `SHIPPED_OVERLAP="union"`.
- `entity_field_resolution` plant/year conflicts are expected (multi-plant jobs, multi-year jobs) and
  do NOT drive `needs_review`.
- Re-running is idempotent given unchanged sources; `data/resolutions.csv` carries decisions forward.

## Current scale (2026-Q2 run)

~93,137 shipped pieces (base 20,987 / rehab 3,304 / non-base 68,750 / unknown 96), 3,137 job codes,
609 jobs needing location review, 188 jobs released-but-not-shipped.

## Pending work (as of 2026-06-15)

### County review — partially complete

`output/review/04_county.xlsx` has 154 conflict rows; 7 confirmed, 147 remaining.

- **DIO** — "FREEDOM" is a town not a county; needs actual county lookup (likely Outagamie County WI).
- **BLL** — DESCHUTES (jobcode/registry) vs MARION OR (149 ERP lines); needs manual project-location lookup.
- **19 bucket codes** (BV, BX, BZ, CA–CN, MH, ST) — catch-all ERP codes mixing multiple projects; county unresolvable without sub-job attribution; flagged with notes.
- When shipping map uses county data, note that **147 of 154 county conflicts are unconfirmed** — those jobs will have lower-confidence county assignments.

### jl_mdrive_reader.py extractor improvements (M: drive required)

M: drive was not accessible on 2026-06-15. Deferred to next session with M: drive access:

- Extract `NCF county` field from New Customer Forms
- Preserve folder project name (currently overwritten by parsed name)
- Cross-validate job code from folder name vs BOM job code

Plan notes: `C:\Users\JohnLeitzke\.claude\plans\c-users-johnleitzke-claude-plans-read-my-enumerated-knuth.md`
