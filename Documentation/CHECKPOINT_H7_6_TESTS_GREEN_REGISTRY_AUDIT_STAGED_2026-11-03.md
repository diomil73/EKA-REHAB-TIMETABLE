# H7 focused fixes green and registry comparison staged | 2026-11-03

## Verified operator source gate

```text
6 passed in 1.78s
```

This covers `Tests/test_h7_patient_mode_switch_and_toolbar_fix.py` and `Tests/test_h7_patient_coverage_audit.py`, after commit `551c75b` for explicitly opaque black patient statistics navigation text. It is **source GREEN**, not native Excel UI proof.

## Current evidence and open questions

- BD integrated structural audit GREEN, but native click on Overall had once produced a generic error; subsequent user clicks succeeded. Treat as intermittent, not root-caused. Unmerge-before-clear fix and additional error details are staged for a future disposable workbook, not installed on BD.
- Both statistics buttons appear misaligned; one had white text. Normalization source and regression tests now GREEN; native verification outstanding.
- BD read-only patient history coverage: 80 finalized sessions, 77 patients, 77 performed, 3 absences. All four period reports reconcile with the same sessions. Patient IDs missing in history source: 02, 06, 15, 19, 20, 21, 24, 25, 30, 34, 39, 41, 50, 64, 68, 69, 73, 75, 82, 84.
- User specifically reports absent 15, 18–22 in daily/weekly views. 15,19,20,21 are absent from CURRENT_FINAL session data; 18 and 22 **are present in session source** and require UI visibility confirmation.
- **Do not fabricate sessions** for registered patients without finalized activities. Distinguish activity report vs a complete registry/zero-activity report.

## New stage: registry-vs-history read-only audit

- `Python/tools/audit_h7_registry_vs_history.py`, commit `77083a2`: compares `PATIENTS` registry, H5 CURRENT_FINAL session history and daily/weekly H7 report patient identifiers. Lists explicit YES/NO for IDs 15,18,19,20,21,22, plus registry numeric IDs without sessions and source IDs absent from registry.
- `Tests/test_h7_registry_vs_history_audit.py`, commit `71f2122`: three unit tests, **not yet operator-proven**.

Next source-and-read-only gate, with Excel closed:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_h7_registry_vs_history_audit.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: registry audit tests failed" }
python .\Python\tools\audit_h7_registry_vs_history.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
```

Do not rerun the already GREEN six tests. If the registry auditor says `H7 REGISTRY AUDIT NEEDS REVIEW`, preserve output: this may be an expected mismatch due to registry state, but is not permission to mutate historic data. If it says `H7 REGISTRY AUDIT COMPLETE`, compare presence states and decide separately whether the operator needs zero-activity registered patients in a different explicit view.

All prior frozen smokes unchanged, including AZ/BB, BC, BD. Native H7 full acceptance pending. Final worksheet protection remains release blocker.
