# H7 BD pre-existing file: safety stop, do not overwrite | 2026-11-03

## Operator report

```text
19 passed in 2.73s
SAFETY STOP: H7IntegratedPreparationError: Destination already exists: refusing overwrite
```

The `prepare_h7_integrated_ui_smoke.py` refused to overwrite `Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm`. **This is intended safety behavior.** It only proves a destination exists; it does not establish whether the earlier installer completed or whether the workbook is safe to open.

## New read-only diagnostic

`Python/tools/inspect_h7_integrated_ui.py`, commit `cbf64d5`: checks MASTER Save Day placement/binding, no old DAILY_INPUT close/save or patient stats button, adjacent therapist/patient navigation on THERAPIST_DAILY, all three VBA modules, visible report sheets and four veryHidden history sheets. Opens ReadOnly=True, closes SaveChanges=False and never invokes installation or saves. Returns nonzero if any required item is missing.

With Excel closed:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python .\Python\tools\inspect_h7_integrated_ui.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
```

If **H7 INTEGRATED AUDIT OK**, this confirms persisted installation structure only; then focused native visual clicks on BD are warranted, still not automatically GREEN/FROZEN. If **AUDIT INCOMPLETE**, preserve BD as diagnostic/FROZEN and do not repair or rerun installers; use a fresh suffix BE after a targeted correction or investigation. If the audit itself errors due to Excel COM, distinguish environmental failure from a workbook defect and do not label BD as broken from that alone.

Frozen earlier artifacts: AZ and BB GREEN; BC provisional/frozen with user UX findings; production immutable. Do not repeat passed 19 source tests absent new source changes. Final worksheet protection still release-blocking.
