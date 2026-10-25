# DAILY_INPUT → Operational timetable

This workflow remains the authoritative daily-input business path, but its old production transaction mechanism is being retired.

## Locked production rule

`DAILY_INPUT` must not close and reopen the workbook in normal use.

The final production interaction is:

1. enter absences/cancellations in `DAILY_INPUT`;
2. request application;
3. validate the saved current state and business rules;
4. compute the dated operational changes;
5. write the verified changes into the already-open workbook;
6. read back the affected facts/projection semantically;
7. save the same workbook;
8. remain open and route the user to the appropriate next operational screen.

The old preview -> authoritative file replacement -> close/reopen path in `Python/rehab_excel/daily_input_action_vba.py` is transitional migration debt and is not an acceptable production endpoint.

## Development/smoke previews

Developer preview tools may still create disposable copies such as `DAILY_INPUT_APPLIED_PREVIEW.xlsm` for verification. That behavior is test scaffolding only and must not define the normal user experience.

## Business rules

- Selecting a patient/therapist name activates the row.
- `Όλη ημέρα = ΝΑΙ` ignores time fields.
- A patient timeslot affects only that dated occurrence.
- A therapist row with only `Από` means exactly that timeslot. Users do not need to type artificial ranges such as `08:30-08:31`.
- Patient status is the absence/cancellation reason. `ΠΑΡΩΝ` does not create an absence.
- Pair/group cells are rendered patient-centrically, so only the affected patient line is changed.
- The base schedule remains the source of truth.
- Daily facts remain overlays and must not destructively rewrite recurring/base scheduling.
