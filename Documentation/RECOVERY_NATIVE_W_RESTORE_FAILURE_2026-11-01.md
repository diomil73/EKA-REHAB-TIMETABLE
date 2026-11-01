# Recovery native W restore failure — 2026-11-01

## Native result

Fresh canonical workbook:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_W.xlsm`

Scenario:

- DAILY_INPUT date `2026-10-30`
- patient `ΖΑΛΟΚΩΣΤΑΣ`
- all day `ΝΑΙ`
- reason/status `διακομισμένος`
- patient absence Apply succeeds natively
- restore is then attempted once

Post-run inspection proves the absence Apply remains committed:

- `THERAPIST_DAILY!B7 = ΖΑΛΟΚΩΣΤΑΣ`
- `B7` strike = `True`
- muted/gray presentation remains
- infectious yellow fill remains
- `MASTER_SCHEDULE!D8` strike = `True`, red
- `MASTER_SCHEDULE!M8 = διακομισμένος`, red
- ledger row 2 remains `ACTIVE`
- `RestoredAt` remains empty
- display row = 23

Restore diagnostic:

- timestamp `2026-11-01 18:44:48`
- phase `apply restore payload`
- error `Restore patch apply failed`

Therefore the restore transaction did not commit. W is frozen as failed restore evidence and must not be retried or repaired in place.

## New formula-corruption evidence

After the failed restore/rollback, MASTER formulas show external-workbook qualifiers such as:

`=[1]PATIENT_PLANNER!F4`

instead of the internal reference:

`=PATIENT_PLANNER!F4`

This is a separate rollback bug, not evidence that the patient-centric engine failed.

Canonical `RestorePatchesFromRollback` currently performs cross-workbook Excel copy:

`rollbackBook.Worksheets(sheetName).Range(cellAddress).Copy Destination:=ThisWorkbook.Worksheets(sheetName).Range(cellAddress)`

Excel can rewrite internal formula references when a formula cell is copied between different workbooks. This explains the `[1]PATIENT_PLANNER` qualifier observed only after rollback.

## Restore D8 root-cause finding

`prepare_daily_change_restore_in_place.py` clean-day restore currently handles `master_daily_non_active_overlay` by reading the current MASTER cell formula as text and creating a normal value patch:

- expected patch value becomes the formula string, e.g. `=IF(PATIENT_PLANNER!F4=...)`
- VBA `ApplyPatch` writes that string to the formula cell
- VBA `VerifyPatch` verifies value intent with `CStr(cell.Value)`
- for a formula cell, `cell.Value` is the calculated result, not the formula text

Therefore a value-style read-back comparison is intrinsically invalid for this restore patch. The correct restore contract is presentation-only for the non-active MASTER treatment cell: preserve the existing formula and clear only daily overlay presentation (red/strike).

## Bounded recovery plan

Do not modify replacement/global solver or patient-centric rendering.

Prepare a source fix with regression tests that does only:

1. make clean-day MASTER non-active restore presentation-only,
2. add explicit `master_daily_restore_overlay` presentation apply/verify behavior to clear red/strike without rewriting the formula,
3. make rollback formula-safe so cross-workbook copy cannot leave `[1]...` external references,
4. preserve all existing apply/business behavior.

Before another native workbook, run focused Python regression tests. W remains frozen.

## Additional forensic check available without rerun

Because restore internally calls `ApplyValidatedDailyInputPayload`, `%TEMP%\eka_daily_input_last_error.log` may contain the exact inner patch index/sheet/cell from the W restore attempt. Read it before any new native run if still present.
