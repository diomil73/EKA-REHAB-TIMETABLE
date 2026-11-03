# H7 integrated toolbar source gate GREEN | 2026-11-03

User-confirmed:
```text
19 passed in 2.73s
```

Source gate includes `Tests/test_patient_history_ui_h7.py`, `Tests/test_patient_history_reports_h7.py`, and `Tests/test_h7_integrated_toolbar_and_patient_order.py`. This proves Python/source tests only.

**BD native installer outcome has NOT YET been supplied.** Do not assume BD exists, was built, or has been opened.

Next gate after closing Excel:
```powershell
python .\Python\tools\prepare_h7_integrated_ui_smoke.py --output ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BD installation failed" }
python .\Python\tools\inspect_unified_day_save_controls.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
```

If BD already exists, do **not** rerun installer or overwrite. Instead provide existing installer output and inspect read-only. Expected installer `H7 INTEGRATED UI INSTALL OK` after separate H6/H4/H7 successes. Expected MASTER only one Save Day, DAILY_INPUT no close/amend/save or patient stats button, THERAPIST_DAILY patient/therapist stats buttons aligned, report buttons functional, patient ordering natural and display 01–09. Source AZ remains frozen; BC is frozen provisional observation candidate; BB is H6 final layout proof.

If install or live UI fails, freeze BD failed and make a new uniquely suffixed candidate after source fix. If native UI confirmed, mark BD green/frozen. No new v1→v2 save needed. Production immutable. Worksheet edit protection remains release blocker.
