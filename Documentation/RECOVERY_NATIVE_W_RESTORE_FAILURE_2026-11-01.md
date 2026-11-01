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

## Confirmed inner failure

The preserved inner DAILY_INPUT diagnostic from W confirms the exact restore failure point:

- timestamp `2026-11-01 18:44:35`
- phase `verify patch`
- patch index `2`
- sheet `MASTER_SCHEDULE`
- cell `D8`
- error `DAILY_INPUT read-back mismatch`

This confirms the D8 formula/value restore contract as the root cause of the W restore failure.

## Source fix and green gate

Canonical source fix commit:

`70f1564` — `Fix DAILY_INPUT restore formula contract`

Focused regression gate after the bounded source changes:

`17 passed in 4.43s`

The source fix does only:

1. clean-day MASTER non-active restore is presentation-only,
2. restore source tags clear red/strike without rewriting the formula,
3. rollback explicitly restores the original formula after cross-workbook formatting copy,
4. existing apply/business behavior remains unchanged.

## Native X success

Fresh canonical workbook:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_X.xlsm`

Same patient-absence scenario was applied once, then restored once.

Observed native behavior:

- patient removed correctly from MASTER during the active absence,
- patient removed correctly from therapist schedule during the active absence,
- Restore returned the patient correctly everywhere,
- DAILY_INPUT source row was cleared,
- `THERAPIST_DAILY!B7 = ΖΑΛΟΚΩΣΤΑΣ`, strike cleared, normal black text restored,
- `MASTER_SCHEDULE!D8` restored with normal internal `PATIENT_PLANNER` formula and no `[1]` qualifier,
- `MASTER_SCHEDULE!D8` strike cleared and black presentation restored,
- `MASTER_SCHEDULE!M8 = παρών`, normal black presentation restored,
- ledger row status = `RESTORED`,
- `RestoredAt = 2026-11-01 19:06:15`.

The restore error log and inner DAILY_INPUT error log displayed after X still contain the older W timestamps (`18:44:48` and `18:44:35`). They are stale evidence from W, not X failures. X restore completed successfully after those timestamps.

Therefore X is the first confirmed native green patient-absence apply + restore cycle on the recovered canonical path.

## UX follow-ups, not recovery blockers

Two non-blocking improvements remain:

1. During the active absence state, the visible status text `διακομισμένος` may still retain strike-through. This is cosmetic and does not block operational correctness.
2. The ledger currently records `Κανονικό πρόγραμμα` as `FromText` and does not show the practical origin slot/therapist/time for an all-day patient absence. For operator clarity and future history/statistics, a later usability enhancement should enrich the change description with the original therapist and time information when available, even when the patient did not transfer elsewhere.

These are post-recovery usability/history enhancements. They must not reopen the now-green restore transaction or replacement solver.

## Decision after X

The native patient-absence transaction gate is now green:

- Apply: green
- authoritative persistence: green
- ledger persistence: green
- Restore: green
- formula-safe rollback contract: fixed

The recovery blocker is closed for this scenario. Resume the planned close-day/history/statistics track, while keeping the two UX items above in the backlog.
