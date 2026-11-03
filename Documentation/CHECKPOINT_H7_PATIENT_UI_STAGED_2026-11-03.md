# H7 patient history UI staging | 2026-11-03

## GREEN proofs
- User supplied `22 passed in 5.12s` on H7 patient/report/aggregation suites.
- Real read-only patient monthly TSV from frozen AZ: `Scheduled: 80`, `Performed: 77`, `PatientAbsences: 3`, `Cancellations: 0`, `Replacements: 0`, `Patient report rows: 77`.
- Distinguish **77 patient summary rows** from **80 session records**: monthly grouping is by patient. This is expected, not a discrepancy. This proof does not yet validate the in-workbook UI.
- H7 general October report additionally returned `ok:true`, `finalized_days:1` and `workbook_unchanged:true`.

## STAGED source implementation, NOT native-tested
- `Python/rehab_excel/patient_history_report_vba.py` commits `4361d11`, `74a7d66`: same-session Excel patient history display; four report modes, protected report surface, back-to-MASTER; output from read-only H7 TSV exporter.
- `Python/tools/install_patient_history_ui.py` commit `5de4690`: native VBA + screen installer and read-only reopen confirmation; requires technical history sheets to remain veryHidden.
- `Tests/test_patient_history_ui_h7.py` commit `e2453fa`: source regression, not yet executed by operator.
- Report entry button: at bottom of DAILY_INPUT, `Στατιστικά ασθενών`. Does not touch the MASTER-only `Αποθήκευση ημέρας` control.
- It reuses the H4 therapist report's UI pattern but **does not modify H4**.

## Next source gate
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_patient_history_ui_h7.py Tests/test_patient_history_reports_h7.py -q
```
Expected **12 passed** (5 new UI source tests + 7 existing report tests), subject to actual operator proof.

## Future native gate after source green
Use a unique fresh successor (e.g. `BC`) created by **copying** frozen AZ to a new path, with Excel closed and SHA-256 equality verification before installation. **Do not run the blank H5 builder** because it lacks finalized history needed to test reports. Do not alter AZ. Install H7 patient history UI on fresh copy, then read-only audit and native click through daily/weekly/monthly/overall/back. Verify actual patient totals and button persistence, historical technical sheets still veryHidden, and MASTER Save Day unaffected. If fail, freeze BC and use next unique artifact for code fixes. Do not yet request native test until source gate green.

## Release blocker
Fine-grained worksheet protection must be addressed before production. Preserve all green/frozen artifacts and immutable `Rehab_Center_System_v27_1.xlsm`; defer cosmetic underline.
