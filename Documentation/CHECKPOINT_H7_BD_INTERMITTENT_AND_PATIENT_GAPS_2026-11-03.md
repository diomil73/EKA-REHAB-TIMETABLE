# H7 BD native report observations and unresolved diagnosis | 2026-11-03

## Operator findings
- BD structural read-only audit previously GREEN: H6 Save Day placement, two H4/H7 entries, three VBA modules, four technical history sheets veryHidden.
- **Native click on patient Σ`υνολικά` produced generic error**: "Δεν ήταν δυνατή η προβολή του ιστορικού ασθενών. Δεν τροποποιήθηκε το ιστορικό."
- **Subsequent repeated button clicks worked**, per user. Treat as *intermittent / not yet root-caused*. Do not claim permanent failure or native full GREEN.
- Companion statistics buttons look crooked and one has white text.
- In daily and weekly reports, operator noticed missing patient numbers, including 15 and others. Determine whether patients are in CURRENT_FINAL rows and omitted by report, or have no finalized sessions for that period. Never fabricate sessions to make numbering continuous.

## Status
BD is **NOT accepted** as integrated native GREEN. Preserve it without installer reruns or edits. Prior success of structure does not prove all report paths. Source AZ/BB remain GREEN/FROZEN, BC observational/frozen.

## Focused source fixes staged, not native proven
- `Python/rehab_excel/patient_history_report_vba.py`: unmerge header before clearing on mode refresh (possible Excel 1004), preserve error number/description for subsequent diagnostic, unify buttons' geometry and black font. Commits `964e784`, `3cc0da2`.
- `Tests/test_h7_patient_mode_switch_and_toolbar_fix.py`: 3 targeted source regressions, commit `85fe16c`.
- `Python/tools/audit_h7_patient_report_coverage.py`: read-only coverage audit comparing daily/weekly/monthly/overall report patient IDs and totals against CURRENT_FINAL H5 rows; displays numeric gaps as *informational*. Commit `ec11f3f`.
- `Tests/test_h7_patient_coverage_audit.py` commit `f3ba7e4`, 3 tests.
- These updates are **source-only and unverified by local Windows pytest** at checkpoint; no fixes installed in BD.

## Next tests, no new Excel install yet
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_h7_patient_mode_switch_and_toolbar_fix.py Tests/test_h7_patient_coverage_audit.py -q
```
Expected six source tests; do not claim passed before operator output.

Read-only coverage inspection (with Excel closed):
```powershell
python .\Python\tools\audit_h7_patient_report_coverage.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
```
This checks all four report modes against CURRENT_FINAL and indicates whether patient 15 is present in source. If coverage is complete but Excel shows absent rows, investigate UI rendering/window/filters. If the source lacks ID 15, inspect registry/plan and distinguish absence of a session from an actual reporting defect. If a report/export fails, capture exact console error; do not edit BD.

## Safety
No changes to production or previous green smokes. When targeted source tests and read-only coverage are understood, integrate corrected source into a new uniquely named smoke (BE), not BD. Full screen visuals still require native verification. Final worksheet edit-protection hardening remains release blocker.
