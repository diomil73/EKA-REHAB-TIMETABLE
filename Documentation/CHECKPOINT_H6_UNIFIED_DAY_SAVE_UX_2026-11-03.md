# H6 unified day-save UX | staged 2026-11-03

## Goal
Replace the competing DAILY_INPUT "Κλείσιμο ημέρας" and "Αποθήκευση ημέρας" buttons with a single contextual "Αποθήκευση ημέρας" route. Provide a second shortcut on MASTER after the existing "Εφαρμογή & Συνέχεια" navigation returns there.

## Implementation (source staged, NOT native-proven)
- Installer: `Python/tools/install_unified_day_save_ui.py` commit a09e974.
- Regression: `Tests/test_unified_day_save_ui.py` commit e47d775.
- New VBA `modUnifiedDaySave.SaveOperationalDay` reads DAILY_INPUT!B2 and checks FINALIZED_DAY_SNAPSHOTS by service date and CURRENT_FINAL status.
- If zero current snapshot and no history for the date, delegates to existing `CloseOperationalDay`.
- If exactly one current snapshot, delegates to existing `StartClosedDayAmendment`.
- Ambiguous/inconsistent history is a safety stop. No historical rewrite by dispatcher.
- Installer removes old two DAILY_INPUT shapes and puts one contextual shape on DAILY_INPUT and a shortcut on MASTER.
- Reopens saved workbook read-only and validates persisted macro and controls.

## Next native proof
Run source regressions on Windows, build a fresh preview **AZ** from immutable production, run `install_excel_time_readback_fix.py`, then `install_unified_day_save_ui.py` and verify their OK markers. Open AZ and test first click from DAILY_INPUT produces v1; read-only close-day inspector must be green. Apply a real day change, observe return to MASTER and use its new Save Day button, verify v1->v2 with amendment inspector. Do not reuse AX or AY.

## Safety statuses
- AX and AY remain GREEN/FROZEN H5 proofs.
- H6 is staged and NOT yet validated on native Excel.
- Failed AZ remains failed/frozen if any native failure; use new suffix for subsequent experiments.
- Keep the spelling underline for 'διακομισμένος' as deferred/no-action.
