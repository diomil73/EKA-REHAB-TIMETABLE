# H7 BE screenshot: unexpected Refresh button and toolbar placement | 2026-11-04

## Native user evidence

Operator shared an Excel screenshot of `UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm` on `THERAPIST_DAILY`. Visible issues:
- A conspicuous button labeled `ΑΝΑΝΕΩΣΗ` at the far right of the schedule, detached from the main navigation.
- `Στατιστικά θεραπευτών` and `Στατιστικά ασθενών` placed beneath the restoration controls instead of the main toolbar row; geometry and spacing are not acceptable.
- `← MASTER` and `Save & Exit` are visibly paired at lower left on the main toolbar row. The two statistics buttons should be aligned to that row and kept clear of restoration controls and table.

Screenshot itself is **not evidence** that Refresh is safe to delete or that a full functional H7 test passed. No assertion of BE GREEN/FROZEN. Do not mutate BE or any previous smoke.

## Diagnostics committed

`Python/tools/audit_h7_therapist_daily_shapes.py` commit `bcd60a7` enumerates all `THERAPIST_DAILY` shapes read-only, with caption/name/OnAction/position/dimensions/font color. This will determine which macro (if any) is behind the anomalous `ΑΝΑΝΕΩΣΗ` button and confirm exact toolbar coordinates.

With Excel closed:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python .\Python\tools\audit_h7_therapist_daily_shapes.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm"
```
Use this output to determine whether refresh is obsolete or essential before deleting any shape. If refresh is a legacy unused control, remove it in next installer source and add regression + persisted audit; do not remove blindly.

## Proposed follow-up

Align H4/H7 entry buttons with the existing `ekaToolbar_ExitApp` on `THERAPIST_DAILY` (same top row, sequential to Save & Exit), not with `UsedRange` last row. Keep black text and nonoverlapping geometry. Verify no overlap with cell labels/restore actions in the screenshot. Source changes and tests first, then a **new uniquely suffixed BF** candidate from AZ, not a repair of BE. Ensure release-related sheet protection remains pending.

## Subsequent targeted source changes, unverified

- `Python/rehab_excel/patient_history_report_vba.py` commit `3acdf55`: in next fresh workbook, align H4 and H7 statistics buttons directly after `THERAPIST_DAILY` `ekaToolbar_ExitApp` in the main toolbar row, instead of the dynamic bottom of `UsedRange`.
- `Python/tools/install_patient_history_ui.py` commit `9bd18b5`: persistence audit fails when the two buttons are not aligned with `Save & Exit`.
- `Python/tools/inspect_h7_integrated_ui.py` commit `dad5487`: stricter independent read-only audit for toolbar-row placement.
- `Tests/test_h7_integrated_toolbar_and_patient_order.py` commit `306a29c`: updated source expectations.

**Do not claim source GREEN until running**:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_h7_integrated_toolbar_and_patient_order.py Tests/test_patient_history_ui_h7.py Tests/test_h7_patient_mode_switch_and_toolbar_fix.py -q
```
Expected **15 passed** (7 + 5 + 3), not operator-confirmed.

**Crucial:** these changes do NOT yet suppress or remove the mysterious `ΑΝΑΝΕΩΣΗ` control. Await read-only shape enumeration of BE, then trace and retain/remove it according to function, before preparing BF. BE must not be modified or reused.


## Follow-up operator evidence: exact BE shape inventory

Operator read-only audit of `UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm`:
- `Button 1`, caption `ΑΝΑΝΕΩΣΗ`, OnAction `UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm!RefreshAll`, x=1228.7, y=30.3, 91.6x29.4. This is a generic legacy shape; **do not remove until the macro body and workbook refresh dependencies are inspected**.
- Main nav `ekaToolbar_BackToMaster`: x=4, y=563.6, 82x24.
- `ekaToolbar_ExitApp`: x=92, y=563.6, 70x24.
- Restore controls y=604..672.
- Statistics buttons: `ekaToolbar_TherapistHistory` x=164, y=664.4, 146x26 and `ekaToolbar_PatientHistory` x=318, y=664.4, 146x26.
Thus BE confirms the statistics controls are still in the lower restore-controls region, while the main navigation row is y=563.6. This BE was created **before** the new source positioning patch; do not repair it. The patch for the next fresh workbook places both stats buttons at y=563.6 after Save & Exit.

## New read-only diagnostic
- `Python/tools/audit_h7_refresh_control.py`, commit `4034bb3`: inspects VBA component code for `RefreshAll` procedure/references and counts workbook connections, pivot caches, query tables and pivot tables. It does not call refresh, click controls, or save. This will help determine whether legacy button is needed. Operator has not yet run it.

## Next combined gate (Excel closed)
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_h7_history_previous_screen_navigation.py Tests/test_therapist_history_ui.py Tests/test_patient_history_ui_h7.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Navigation tests failed" }
python .\Python\tools\audit_h7_refresh_control.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm"
```
Expect 18 passed, pending actual operator result. Refresh audit should show VBA procedure source and refresh dependencies. If VBA project access is denied, use Excel Trust Access settings only if already authorized; do not modify workbook. Based on audit, retain or remove legacy RefreshAll in a future fresh BF. Never infer safety from its visual appearance alone.
