# Recovery native S build checkpoint — 2026-11-01

## Status

Fresh canonical preview S built successfully from the normal unified builder. No post-build recovery wrapper was used.

Artifact:
- `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_S.xlsm`
- size: 472583 bytes
- last write: 2026-11-01 11:51:06 local
- SHA-256 prefix observed: `E7AEAB0BEDE823CB78312C9437331686212EF730A03...`

Build evidence:
- registration menu + forms: OK
- patient registration bridge: OK
- patient edit flow: OK
- outpatient schedule form: OK
- outpatient conflict guard: OK
- DAILY_INPUT sheet: OK
- MASTER projection refresh: OK (93 patients)
- final user navigation: OK
- MASTER visual baseline: OK
- unified preview: OK
- source unchanged: True
- VBA project present: True
- DAILY_INPUT sheet present: True
- application contract verified: 24 components

Canonical source gate immediately before build:
- 71 passed
- 2 known non-functional Python ZipFile cleanup warnings

## Scope of the S native smoke

Open only S and test one supported patient absence followed by one single-row restore.

Required observations after DAILY_INPUT Apply:
1. MASTER treatment remains visible.
2. MASTER treatment is red and struck through.
3. MASTER status/reason (for example `διακομισμένος`) is red and NOT struck through.
4. THERAPIST_DAILY existing operational information is preserved, including origin therapist/time and destination therapist details.
5. Daily change ledger row appears without overwriting operational content.

Required observations after single-row restore:
1. Exact DAILY_INPUT A:E source row is cleared.
2. THERAPIST_DAILY returns to recurring/base display.
3. MASTER treatment returns to normal.
4. MASTER status returns to base state (`παρών` for current specimen).
5. Ledger row remains visible as `Επαναφέρθηκε`.
6. No VBA compile/runtime highlight.
7. Same workbook remains open.

Do not run `RestoreAllDailyChanges`, therapist absence restore, or `CloseOperationalDay` on S.

## Recovery architecture status

Post-build recovery wrapper tooling from N through R is retired and retained only as forensic evidence. Current recovery line uses canonical sources plus `Python/tools/build_unified_preview.py` only.
