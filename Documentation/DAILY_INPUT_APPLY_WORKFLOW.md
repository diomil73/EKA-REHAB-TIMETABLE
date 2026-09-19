# DAILY_INPUT → Operational timetable

This step closes the first end-to-end daily workflow.

1. Create `DAILY_INPUT_PREVIEW.xlsm`.
2. Enter absences/cancellations in `DAILY_INPUT`.
3. Save and close Excel.
4. Run `Python/tools/apply_daily_input.py --overwrite`.
5. Python reads the saved `DAILY_INPUT` rows and creates `DAILY_INPUT_APPLIED_PREVIEW.xlsm`.
6. The input workbook is not modified by the apply step.

Rules:

- Selecting a patient/therapist name activates the row.
- `Όλη ημέρα = ΝΑΙ` ignores time fields.
- A patient timeslot affects only that dated occurrence.
- A therapist row with only `Από` means exactly that timeslot. Users do not need to type artificial ranges such as `08:30-08:31`.
- Patient status is the absence/cancellation reason. `ΠΑΡΩΝ` does not create an absence.
- Pair/group cells are rendered patient-centrically, so only the affected patient line is changed.
- The base schedule remains the source of truth.
