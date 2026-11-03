# H7 BD coverage read-only GREEN; targeted style regression fixed | 2026-11-03

## Actual user output

Targeted source pytest: **5 passed, 1 failed in 4.85s**. Failure exact assertion: `shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0` absent from `install_patient_report_components` (patient navigation button). Operator's PowerShell safety stop executed. The historical coverage audit was subsequently run independently, read-only.

Coverage audit on frozen BD:
- 2026-10-30: **80 CURRENT_FINAL sessions, 77 distinct patients, 77 performed, 3 absences, 0 cancellations**.
- Numeric IDs absent in CURRENT_FINAL source: **02, 06, 15, 19, 20, 21, 24, 25, 30, 34, 39, 41, 50, 64, 68, 69, 73, 75, 82, 84**.
- **15, 19, 20, 21** are therefore missing from the source itself. **18 and 22 are *not* in the absent set**, so must investigate view visibility/ID formatting if the user cannot see them.
- All four modes daily, weekly, monthly, overall: **77 report rows, 80 scheduled, 77 performed, 3 absences, 0 cancellations, 0 replacements**.
- Audit marker: `H7 PATIENT COVERAGE AUDIT OK; no workbook changes`.

**Interpretation:** No demonstrated loss of finalized sessions by the Python summary. Source ID gaps might represent patients without a finalized session on that date; must compare patient registry/schedule before deciding whether to show all patients with zero activity. Current report is an activity report, not necessarily a complete registry roll. User has asked to consider all numbers, especially 15 and 18–22. Never fabricate visits.

## Fixed source

- `Python/rehab_excel/patient_history_report_vba.py`: append explicit opaque black font on new patient statistics shape, commit **`551c75b`**. This addresses the single test failure; **not yet rerun**.
- `Python/tools/audit_h7_patient_report_coverage.py`: print YES/NO for 15, 18, 19, 20, 21, 22 in source, commit **`41bf893`**. Existing full audit already provides absent IDs; no need to rerun it just for this decoration.
- Earlier staged unmerge-before-clear on report-mode switch + diagnostic error number, and matched therapist/patient button geometry, remain **not native-proven**. Previous intermittent `Συνολικά` error could not yet be root-caused definitively; unmerge addresses a plausible Excel merge-related source of error.

## Next gate

```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_h7_patient_mode_switch_and_toolbar_fix.py Tests/test_h7_patient_coverage_audit.py -q
```

Expect **6 passed**, not yet operator-confirmed. Then investigate **registry membership / original schedule** separately from CURRENT_FINAL activity. Keep BD as diagnostic/FROZEN; no reinstall or edits. New uniquely named BE only after questions about expected report population and necessary source corrections are resolved.

## All prior safety

AX/AY/AZ, BB and AQ GREEN/FROZEN; BC provisional/frozen; BD diagnostic/frozen. Production `Rehab_Center_System_v27_1.xlsm` immutable. Final worksheet editing protection is still release-blocking.
