# Recovery native M build — 2026-10-31

## Artifact

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_M.xlsm`

## Build evidence supplied by operator

Canonical unified preview build completed successfully with all reported stages green:
- registration menu + forms,
- patient registration bridge,
- patient edit flow,
- outpatient schedule form,
- outpatient conflict guard,
- DAILY_INPUT sheet,
- MASTER projection refresh with 93 patients,
- final user navigation,
- MASTER visual baseline.

Build summary:
- `UNIFIED PREVIEW OK`,
- source: `Rehab_Center_System_v27_1.xlsm`,
- preview: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_M.xlsm`,
- DAILY_INPUT date: 2026-10-30,
- production source unchanged: True,
- VBA project present: True,
- DAILY_INPUT present: True,
- application contract verified,
- explicit contract verification: 24 required components,
- artifact SHA-256 begins `C41AA27AEF444EE2C42D0A95C6B9EB126D65ADE55F3...`,
- artifact size: 470921 bytes,
- last-write: 2026-10-31 18:59:55 local.

## Pre-build source gate

The focused gate immediately before M was green:
- 60 passed,
- 2 known Python 3.14 ZipFile cleanup warnings,
- warnings are not functional test failures.

## Changes included since L

M contains the layout-aware/fail-soft MASTER patient-status normalization repair:
- locate the actual MASTER header row using the layout contract rather than assuming row 1,
- find `Κατάσταση` in that row,
- normalize non-empty status cells to `Strikethrough=False`,
- if no status column exists, exit safely rather than failing the whole operational refresh.

Relevant commits:
- `bb785772c68b51bc2adacdf8001fef7e70cc67e5` layout-aware/fail-soft normalization,
- `f03c32215a4f1eb2d8b6aa9a12d77e57b03d98aa` regression for the native L failure,
- `8b796c27897e1a1511756b18eb35f5d395fec759` L refresh-failure checkpoint,
- `cbd3d42174b48dd8a57c35867546e55d054f1c4e` consolidated J-to-M handoff.

## Next native gate

Open only M and run one narrow patient-absence/restore cycle:
1. create one patient absence in DAILY_INPUT and apply it,
2. verify MASTER recurring treatment stays visible and is red + struck,
3. verify MASTER status/reason is red but not struck,
4. verify the daily-change ledger row appears,
5. restore that row,
6. verify DAILY_INPUT source row clears,
7. verify THERAPIST_DAILY and MASTER treatment return to recurring state,
8. verify MASTER status returns to base value (`παρών` for the chosen specimen),
9. verify the ledger row remains visible as restored,
10. verify no central-view refresh error appears and the same M workbook remains open.

Do not run `CloseOperationalDay` yet. Therapist-absence restore remains disabled until its coordinated replacement-fact rollback is implemented and native-proven.
