# H7 patient report presentation preparation | 2026-11-03

## Established GREEN baseline (do not rerun without a reason)

- H2–H6 native acceptance preserved. H5 AX / AY, H6 AZ / BB are GREEN/FROZEN.
- H7 existing source regression: **15 passed in 3.29s**.
- Real historical month report from frozen AZ: October 2026 one finalized day, scheduled=80, performed=77, patient absences=3, cancellations=0, replacements=0; `workbook_unchanged=true`.
- Canonical references: `Documentation/PROJECT_RESUME_START_HERE_2026-11-03.md`, `PROJECT_HANDOFF_EVIDENCE_2026-11-03.md`, `PROJECT_HANDOFF_NEXT_STEPS_2026-11-03.md`.

## This new H7 slice: staged, not yet pytest-verified

- `Python/rehab_excel/patient_history_reports.py`, commit `fcaf91a`: daily, ISO-week, calendar-month and overall summaries per patient; only from `read_current_patient_session_history_rows`. Report counts scheduled / performed (active and replaced) / absence / cancelled / replaced. Reject unresolved and duplicate session identifiers; filter dates inclusive.
- `Python/tools/export_patient_history_ui.py`, commit `54ced24`: read-only UTF-8 BOM TSV output suitable for a later Excel UI, explicit schema, refuse overwriting existing output, reject embedded tab/newline.
- `Tests/test_patient_history_reports_h7.py`, commit `80af4c4`: 7 tests covering all grouping modes, year boundary, exclusion/filtering, duplicate/unresolved safety, schema and non-overwrite.

## Next test gate

```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_patient_history_reports_h7.py Tests/test_patient_statistics_h7.py Tests/test_historical_statistics_h7.py Tests/test_period_statistics_h7.py -q
```

Expected **22 passed** if all new and existing tests pass. Not yet verified; do not assert green before user supplies output.

After green: read-only export against existing frozen AZ to a **new unique temp output file**, verify report values against H7 month totals (patient sums should be 80 scheduled, 77 performed, 3 absences; no double-count). Only then implement patient Excel report VBA/UI reusing the proven H4 therapist report pattern, in a fresh workbook and with native acceptance. Never modify AZ, BB or production.

## Safety / design

- H4 `THERAPIST_HISTORY_REPORT` is already GREEN; do not duplicate or rebuild its current reporting UI.
- H7 patient report is not yet Excel-embedded and has no native acceptance.
- User-facing report must not expose snapshot ids/versions, but keep their precise authoritative filtering internally.
- Preserve technical history veryHidden and protected. No arbitrary sheet unlocking or workbook writes to historical records.
- Final sheet edit-protection hardening remains release-blocking and separately documented.
