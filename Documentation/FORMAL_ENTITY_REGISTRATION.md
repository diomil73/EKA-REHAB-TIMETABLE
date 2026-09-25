# Formal entity registration - validation stage

This stage defines the validation contract that future menu/forms must pass before they write authoritative data.

## Patient

Authoritative source remains `PATIENTS`.

Validated before write:
- non-empty `PatientID`
- unique `PatientID`
- non-empty visible patient name
- optional room must belong to configured `SETTINGS` rooms when a room list is supplied
- optional status must belong to the configured patient-status list when supplied

Duplicate visible patient names are allowed. Identity is based on `PatientID`, not on the display name.

## Therapist

The current workbook provider registry remains `SETTINGS` / `THERAPISTS_FTH` for this stage.

Validated before write:
- non-empty therapist name
- no case-insensitive duplicate therapist name in the current registry

No workbook write is added here yet.

## Student

The domain model already has stable `student_id`, `student_number`, placement dates, supervisor, replacement capability and robotic capability.

Validated before write:
- non-empty stable student ID
- unique student ID
- non-empty display name
- positive and unique student number
- placement end cannot be before placement start
- supervisor, when provided and a therapist registry is supplied, must exist

Duplicate student display names are allowed because identity is kept separately.

## Important boundary

This PR is validation-only. It does not yet invent a new Excel storage location for students and it does not write to the baseline workbook. The next stage can add the actual Excel menu/form writeback once the student authoritative registry location is fixed explicitly.
