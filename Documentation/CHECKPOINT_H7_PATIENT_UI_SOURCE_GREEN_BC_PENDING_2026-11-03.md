# H7 patient UI source GREEN / native BC pending | 2026-11-03

## Last confirmed source gate

User supplied `12 passed in 3.27s` for:
```powershell
python -m pytest Tests/test_patient_history_ui_h7.py Tests/test_patient_history_reports_h7.py -q
```
This validates source-level patient-report UI and exporter regressions. It does **not** yet prove native VBA execution, button positioning, or UI persistence.

## Previously confirmed native/read-only green data

- Frozen AZ month summary: October 2026, one finalized day, 80 scheduled, 77 performed, 3 patient absences, zero cancellations, zero replacements.
- Patient TSV read-only on AZ: 77 monthly per-patient rows summing to 80 scheduled, 77 performed, 3 absences, 0 cancellations and 0 replacements.
- H6 AZ: functional v1→v2; H6 BB: final MASTER-only Save Day layout. H5 AX and AY GREEN. Do not alter these files.
- Production `Rehab_Center_System_v27_1.xlsm` immutable.

## Next native gate: new BC built by byte-for-byte copy

With all Excel windows closed, pull branch then use PowerShell. Confirm BC does not exist, copy AZ to BC and check SHA-256 equality before installing H7 UI. Install only on BC. If any command fails, stop and freeze BC FAILED/FROZEN; never repair or retry it. Installer is `Python/tools/install_patient_history_ui.py`; it performs VBA/module/sheet/entry checks and a read-only reopen check after save. Its success marker is `H7 PATIENT HISTORY UI INSTALL OK`. 
After installer success, open **BC**, visit DAILY_INPUT and click `Στατιστικά ασθενών`. Confirm four modes `Ημερήσια`, `Εβδομαδιαία`, `Μηνιαία`, `Συνολικά` and `← MASTER`, including button persistence when switching modes. Verify at least the monthly visible totals (80 scheduled, 77 performed, 3 absences, 0 cancelled, 0 replaced), technical history stays hidden, MASTER Save Day still visible and correctly placed. The displayed row count is expected to be 77 grouped patient records, not 80 distinct sessions.

## Diagnostic branches
- Copy/hash mismatch: stop, do not install.
- Installer failure: preserve complete console output; do not open or modify BC; next candidate gets unique suffix.
- Successful installer but control absent: inspect BC with read-only COM, do not mutate.
- Native runtime failure or visual defect: freeze BC, capture exact message, stage surgical source fix and fresh unique next artifact.
- All native checks GREEN: mark BC GREEN/FROZEN and checkpoint. Do not rerun day save on BC.
- Final sheet protection remains a mandatory separate release gate; see `Documentation/BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md`.

## Resume references
`Documentation/PROJECT_RESUME_START_HERE_2026-11-03.md`, `PROJECT_HANDOFF_EVIDENCE_2026-11-03.md`, and `PROJECT_HANDOFF_NEXT_STEPS_2026-11-03.md`. This checkpoint updates the H7 UI stage and overrides older "next commands".
