# H7 BD saved report visibility proof; patient-row readability patch | 2026-11-04

## User-confirmed read-only evidence

```text
H7 PATIENT REPORT VISIBILITY READ-ONLY AUDIT
ReadOnly: True
Saved mode label: Μηνιαία σύνοψη ανά ασθενή
AutoFilterMode: False
FilterMode: False
ScrollArea:
Last populated patient row: 82
ID 15: NOT IN SAVED REPORT SHEET
ID 18: row=20 hidden=False height=13.8 text='18'
ID 19: NOT IN SAVED REPORT SHEET
ID 20: NOT IN SAVED REPORT SHEET
ID 21: NOT IN SAVED REPORT SHEET
ID 22: row=21 hidden=False height=13.8 text='22'
H7 PATIENT REPORT VISIBILITY AUDIT COMPLETE; workbook unchanged
```

**Meaning:** the saved BD workbook is in monthly mode, with 77 populated patient rows (worksheet rows 6..82). IDs 18 and 22 are genuinely present and not hidden or filtered. The earlier user report that they visually seemed missing in daily/weekly mode is not directly disproven by inspection of the saved *monthly* mode. Readability is a plausible contributing factor, but this audit cannot prove why they were not noticed in other modes. IDs 15,19,20,21 have no FINAL session in the source, even though they belong to the current PATIENTS registry. The read-only registry-versus-history audit confirmed 93 registry members, 77 historical patient IDs, 80 sessions, and 16 registry members without sessions.

## Source changes staged, no BD mutation

- `Python/rehab_excel/patient_history_report_vba.py` commit `58ccf26`: increases actual patient report rows' height to 21 after rendering; header row unchanged.
- `Tests/test_h7_registry_zero_overlay_and_visibility.py` commit `9ee1cca`: enforces readable row height.
- Prior commits `e25623c`, `f7812e8`, `2c2592d`, `37a6747`, `ab6c53a`: full-current-registry explicit opt-in overlay, 16 zero-activity rows in the AZ-derived one-finalized-day example, no invented sessions, unchanged totals. Not yet operator-tested or native installed.

## Immediate gate

User has not yet supplied the outcome of the previous new zero-overlay source tests. After adding one readability test, run:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_h7_registry_zero_overlay_and_visibility.py Tests/test_patient_history_reports_h7.py Tests/test_h7_patient_mode_switch_and_toolbar_fix.py -q
```
Expected **19 passed** (9 + 7 + 3), pending actual result; avoid rerunning the older green six tests as a separate command.

If source GREEN, create unique BE from frozen AZ via `Python/tools/prepare_h7_integrated_ui_smoke.py`, but only after explicit full source gating. The new Excel UI would display 93 current registered patients including 16 zero-activity, so it must be clearly titled as registry overlay, NOT an immutable historical patient roster. Reconfirm daily/weekly/monthly/overall and 80/77/3 totals, paired button style, actual row visibility, and intermittent Overall failure. BD remains diagnostic/FROZEN, never repair it.

## Caveat

Do not interpret absence of patient IDs in a one-day finalized dataset as a missed therapy or patient absence. No non-finalized dates may be invented. This 21-point height is source proposal until proven native.
