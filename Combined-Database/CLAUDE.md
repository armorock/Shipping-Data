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
- `parse_part_name.py` — parses any-generation Armorock part number into 15 attribute fields
  (`part_type`, `subcategory`, `generation`, `diameter`, `height`, `opening_diameter`, `troughing`,
  `wall_variant`, `section_suffix`, `lid_suffix`, `box_length`, `box_suffix`, `es`, `de`, `de_count`).
  Also exports `build_gen4_name(attrs)` to reassemble a canonical Gen4 name from those attributes.
  Implements the `/item-name-detail-interpreter` + `/part-name-builder` skills in Python.
- `build_pieces_ledger.py` — Layer 2 (`pieces_ledger.csv`). ERP dedup = max-count-per-system per
  (job, part, date). Overlap rule from `sources.SHIPPED_OVERLAP`. Emits reconciliation + quoted_not_shipped.
  Now also writes `gen4_name` (Column E) and 15 `pn_*` attribute columns for every piece.
  Pre-2018 location fallback: captures `Shipping City / Shippings State / ZipCode` from BABY inline
  and uses them when the registry lookup returns None (covers all 14,863 null-job-code pre-2018 pieces).
- `build_reports.py` — rehab/nonbase report, provenance, structure completeness roadmap, review lists.
  Review list columns: `job_code | field | status | current_winner | winner_sources | project_name | [field extras] | src_bom_union | src_dispatch | src_erp_qb | src_erp_fb | src_erp_ns | src_jobcode_db | src_registry | all_observed_values | Confirmed Value | Notes`.
  City extras: `raw_bom | zip_county | county_flag`. Zip extras: `geo_city | geo_state | confirmed_city | confirmed_state | zip_ok` (GeoNames from `Shipping-Map/data/US.txt`).
  Also produces: `04_county.xlsx` (county CONFLICT/GAP rows) and `07_location_matrix.xlsx` (all jobs, all sources, includes AGREE rows + `street_address` column for geocodable addresses).
- `apply_corrections.py` — reads filled review xlsx -> resolutions -> rebuilds registry/ledger/reports.
- `run_quarter.py` — orchestrator. `snapshot_and_diff.py` — per-quarter snapshot + change report.
- `nsaw_export_bases.py` — standalone NSAW data prep; filters `pieces_ledger.csv` to BASE parts with valid A–E 3-letter job codes (plus null job codes for pre-2018 rows), converts all part numbers to Gen4 canonical names, outputs `output/bases_by_job_code.csv`. Re-run before each NSAW.

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

~93,137 shipped pieces (base 20,987 / rehab 3,304 / non-base 68,750 / unknown 96), 3,169 job codes,
705 jobs needing location review.

## Pending work (as of 2026-06-15)

### County review — complete (2026-06-22)

`output/review/04_county.xlsx` had 154 rows; all resolved:

- **130 rows confirmed** — all valid job codes have a confirmed county value
- **24 rows deleted** — QB parse errors (2-char codes, 4-char ANCP, post-E start codes like MH/ST/STA); these had ERP-only data with no real project behind them
- Key corrections: BCQ→GREENVILLE SC, BFR→DAVIS UT, BHK/BIU/BLB→BRUNSWICK NC, BWN→ANDERSON SC, BVZ/BXD/CAS/CDL→MANATEE FL, BPP→CHARLES MD, CDB→DENTON TX, CDF→CHESTER SC, CXH→BRUNSWICK NC, BWV→WEBER UT, CKD→SNOHOMISH WA

### ERP job code cleanup — next step

348 BASE rows in `pieces_ledger.csv` have invalid job codes (2-letter, 4-letter, or starting past E). All originate from ERP sources (`erp_qb`: 325, `erp_ns`: 12, `erp_fb`: 11) — Dispatch is already clean. The NSAW export filters these at output time. To remove them permanently: correct the job codes in BABY.xlsm and re-run the pipeline.

### jl_mdrive_reader.py extractor improvements (M: drive required)

M: drive was not accessible on 2026-06-15. Deferred to next session with M: drive access:

- Extract `NCF county` field from New Customer Forms
- Preserve folder project name (currently overwritten by parsed name)
- Cross-validate job code from folder name vs BOM job code

Plan notes: `C:\Users\JohnLeitzke\.claude\plans\c-users-johnleitzke-claude-plans-read-my-enumerated-knuth.md`
