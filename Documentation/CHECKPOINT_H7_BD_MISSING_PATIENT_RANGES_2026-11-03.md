# H7 BD missing numbered patient rows: expanded operator observation

## New operator evidence
After report UI observations, user reports missing identifiers **15, 18, 19, 20, 21, 22 and others** from daily and weekly patient statistics.

Do **not** assume contiguous numbering means a session was lost. There are three different populations: patient registry, scheduled/finalized patient sessions per period, and report display. Clarify which population is intended only after read-only comparison.

## Immediate diagnostic pending
Use existing `Python/tools/audit_h7_patient_report_coverage.py` on frozen BD with Excel closed. It compares CURRENT_FINAL source patient IDs and totals against daily/weekly/monthly/overall report rows; reports whether numeric patient 15 is in source and lists all numeric gaps in source. It does not currently establish registry membership. Request exact output rather than guessing.

- If CURRENT_FINAL contains any alleged missing ID but report row omits it, fix reporting/export/UI; add targeted source regression.
- If both CURRENT_FINAL and report omit it, compare against patient registry / schedule for target dates. A patient may legitimately have no finalized session in the chosen period; if user needs complete patient roster, implement a separate registry-joined zero-activity view with explicit indicator, never fabricate sessions.
- If CLI report is complete yet Excel view appears missing, inspect AutoFilter, scroll/hidden rows, row visibility, view formatting, and session-state. Preserve frozen BD.
- If source history counts differ from snapshot (80 scheduled, 77 performed, 3 absences for AZ-derived 2026-10-30), safety stop.

The user's earlier note that patient buttons resumed functioning does not erase the intermittent Overall error. White text / crooked paired stats buttons remain a separate UX defect. Source correction staged but not native proven.

## Safety
BD not GREEN/FROZEN full native; preserve as diagnostic, do not repair. AZ and BB GREEN/FROZEN and production workbook immutable. Next corrected workbook requires fresh unique suffix. Final sheet protection remains mandatory before release.
