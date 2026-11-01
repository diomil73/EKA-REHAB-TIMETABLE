# Recovery old-engine isolation proof — 2026-11-01

## Context
Fresh native preview `UNIFIED_CLOSE_DAY_RECOVERY_20261101_U.xlsm` still reproduced the same native symptom set seen in S/T: no visible progress in the operational projections and the in-place DAILY_INPUT transaction failed during MASTER read-back.

## Isolation experiment
The same U workbook was used as input to the legacy/copy-based DAILY_INPUT application path:

`Python/tools/apply_daily_input.py`

Output:

`Excel/previews/DIAGNOSTIC_OLD_ENGINE_FROM_U_20261101.xlsm`

The source U workbook hash was checked before and after the diagnostic run.

## Result
The legacy/copy-based path completed successfully:

- `DAILY INPUT APPLIED PREVIEW OK`
- Date: `2026-10-30`
- Therapist rows read: `0`
- Patient rows read: `1`
- Operational absences: `1`
- Session cancellations: `0`
- Changed sessions: `1`
- Binding: `patient_absent: planner:4:ΦΘ@2026-10-30 [B7]`
- Input unchanged: `True`
- VBA preserved: `True`
- Source U hash unchanged after the run: `True`

## Architectural conclusion
The patient-centric / DAILY_INPUT planning engine is still operational on the current workbook family. The failure observed in native U is therefore not evidence that the core patient-centric engine, DAILY_INPUT parsing, or copy-based write-plan application is broken.

The current failure boundary is the newer in-session VBA transaction layer that applies and verifies the prepared payload in place. Existing native logs show that the transaction fails during MASTER patch verification (`phase=verify patch`, `patch_index=2`, `MASTER_SCHEDULE`, previously D8/D9 depending on the selected row), after which the transaction rolls back. That rollback explains why downstream THERAPIST_DAILY, status, ledger, and restore effects appear as though nothing happened.

## Recovery rule
Do not alter or replace the patient-centric/replacement engine unless new evidence points there. The next debugging scope is limited to the in-place VBA payload applier / read-back verifier / rollback boundary.

Do not create further alphabetic native builds until the in-place verifier failure is understood with a diagnostic that identifies the exact sub-check failing within `VerifyPatch` / source-tag presentation verification.
