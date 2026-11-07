# H9 Therapist Attendance and Working-Day Historical Evidence

Date: 2026-11-08. Branch: `work/history-statistics-rebuild`.

## Accepted starting point

The H8 **BL** workbook is native-audit and operator-manual accepted. It is frozen. Its third therapist panel correctly reports unavailable historical attendance rather than claiming zero absences. The source code for that accepted UI is not changed in this H9 modeling slice.

## Read-only source audit and why more work is necessary

`THERAPIST_DAY_HISTORY` currently persists therapist utilization (slots, sessions, replacements, cancellations, suspensions), **not a verified presence/absence outcome per expected working day**. `FINALIZED_DAY_SNAPSHOTS` holds finalized dates and CURRENT_FINAL version pointers but **cannot establish that a given therapist was required to work** on each date. `PATIENT_SESSION_HISTORY` cannot establish therapist attendance: a therapist may be present with no sessions, absent with replacements, or not scheduled. The existing `DAY_HISTORY_EVENTS` enum currently provides only DAY_FINALIZED, DAY_AMENDMENT_STARTED, and DAY_AMENDMENT_FINALIZED. Thus none of these three sources alone provides a valid absent-day numerator or expected-workday denominator.

## New pure Python domain contract (no live workbook effects)

Added `Python/rehab_core/therapist_attendance_history.py`, which defines an independently sourced **TherapistAttendanceFact** with therapist, date, CURRENT_FINAL snapshot ID, explicit PRESENT/ABSENT/NONWORKING outcome, provenance and required absence reason. `summarize_attendance` accepts an explicit verified expected-working-date calendar and complete facts. It fails closed on missing dates, duplicates, superseded snapshot reference, mixed therapist evidence, unknown provenance, absent reason, or treating NONWORKING as a working day. An explicitly empty expected calendar yields a **None** presence ratio, not a misleading 0% or 100%.

Focused regression file `Tests/test_therapist_attendance_history_contract.py` covers those failure modes and the valid two-day case. **Tests have not yet been executed by the operator.** This contract is not connected to Excel VBA, the history snapshot writer, or a new technical worksheet.

## Required future design and safety gates

1. Specify where the authoritative staff roster/scheduled working dates and approved leave/holiday exclusions are recorded. An operational `THERAPIST_ATTENDANCE` worksheet by itself is current editable state, not finalized versioned historical proof.
2. Define explicit PRESENT/ABSENT reason codes and proof of the daily roster. Differentiate annual leave, public holiday, sick leave, nonworking day and unknown; agree whether approved leave belongs in the requested absence count and denominator. Preserve the decision and audit source.
3. On deliberate `Κλείσιμο ημέρας`, atomically version or append the attendance facts with the therapist/day snapshot; do not update only the chart. Historical amendment must supersede prior facts without deleting audit history, and reports must select CURRENT_FINAL only.
4. Validate the actual calendar (weekends, local/national holidays, shifts, part-time schedules). Never infer absences from missing therapy rows, 0 sessions or days with no finalized history.
5. Once enough genuine facts exist, integrate read-only historical daily/weekly/monthly/overall aggregates into `export_therapist_history_ui.py` and the third native panel. Until then leave the current honest "Μη διαθέσιμο ιστορικό απουσιών / εργασίμων" placeholder.
6. Enforce fresh single-use workbook suffixes and green source+native audit+operator manual verification before changing the BL baseline.

## Next source-only gate

From repository root with Excel closed:

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest Tests/test_therapist_attendance_history_contract.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: attendance contract source gate failed" }
```

If this source gate passes, the model is validated **only at the pure-domain level**. No H9 native workbook is authorized by this test alone. No new output XLSM has been created.
