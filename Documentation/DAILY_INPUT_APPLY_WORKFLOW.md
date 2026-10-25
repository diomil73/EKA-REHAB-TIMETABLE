# DAILY_INPUT → Operational timetable

This workflow is the authoritative daily-input business path.

## Locked production rule

`DAILY_INPUT` must not close and reopen the workbook in normal use.

Production interaction is now:

1. enter absences/cancellations in `DAILY_INPUT`;
2. request application;
3. save and validate the current workbook state;
4. compute the dated operational write plan;
5. serialize only validated cell patches;
6. apply those patches to the already-open workbook;
7. read back value/presentation semantics;
8. save the same workbook;
9. remain open and route to replacements only when needed, otherwise return to MASTER.

`Python/tools/prepare_daily_input_in_place.py` is the read-only preparation layer. It rebuilds the effective daily state and emits the validated `WritePlan` as a UTF-8 payload. It does not create or replace an authoritative workbook copy.

`Python/rehab_excel/daily_input_action_vba.py` applies the payload in-session. It mirrors the native presentation semantics for values, clears, fills, borders, font roles, wrap text, minimum font size and rich-text runs.

Before mutation, VBA creates a temporary `SaveCopyAs` recovery file. This copy is rollback/backup only. The live workbook is never closed or replaced. If any patch or read-back fails, affected cells are restored from the recovery copy and the action reports failure.

The production VBA action contains no `ThisWorkbook.Close`, reopen handoff, `Application.OnTime`, authoritative commit worker, or visible console execution.

## Development/smoke previews

Developer preview tools such as `apply_daily_input.py` may still create disposable copies such as `DAILY_INPUT_APPLIED_PREVIEW.xlsm`. Those are test tools only and do not define normal operator behavior.

## Business rules

- Selecting a patient/therapist name activates the row.
- `Όλη ημέρα = ΝΑΙ` ignores time fields.
- A patient timeslot affects only that dated occurrence.
- A therapist row with only `Από` means exactly that timeslot. Users do not need artificial ranges such as `08:30-08:31`.
- Patient status is the absence/cancellation reason. `ΠΑΡΩΝ` does not create an absence.
- Pair/group cells are rendered patient-centrically, so only the affected patient line is changed.
- The base schedule remains the source of truth.
- Daily facts remain overlays and must not destructively rewrite recurring/base scheduling.
