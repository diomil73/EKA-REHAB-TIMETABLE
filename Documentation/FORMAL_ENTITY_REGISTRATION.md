# Formal entity registration

This area defines the validation and safe preview-write contracts that future menu/forms must pass before authoritative workbook data is changed.

## Patient

Authoritative source remains `PATIENTS`.

Validated before write:
- non-empty `PatientID`
- unique `PatientID`
- non-empty visible patient name
- optional room must belong to configured `SETTINGS` rooms when a room list is supplied
- optional status must belong to the configured patient-status list when supplied

Duplicate visible patient names are allowed. Identity is based on `PatientID`, not on the display name.

A native Excel preview workflow now writes only to a copied `.xlsm`, verifies the new patient by reading the copy back, and proves the source workbook remained unchanged.

## Therapist

The current physiotherapist registry is `SETTINGS` column A, headed `THERAPISTS_FTH`.

Validated before write:
- non-empty therapist name
- no case- or accent-insensitive duplicate therapist name in the current registry

The therapist preview workflow:
- creates a new `.xlsm` copy
- opens only the copy for native Excel COM writeback
- appends the name to the next logical `THERAPISTS_FTH` row
- verifies the resulting therapist registry by reading the copy back
- verifies the other configured SETTINGS value lists are unchanged
- hashes the source before and after the operation

Robotic capability is not persisted in this stage because the workbook has no confirmed authoritative therapist-capability storage location yet. Requests that attempt to persist it stop safely instead of silently dropping the flag.

## Student

The authoritative student registry is a dedicated `STUDENTS` sheet. This is intentionally separate from `SETTINGS`: a student is a multi-field operational entity with identity, placement dates and supervisor, not a one-dimensional settings list.

Authoritative columns A:H are fixed as:

1. `StudentID`
2. `Φοιτητής`
3. `StudentNumber`
4. `PlacementStart`
5. `PlacementEnd`
6. `SupervisorTherapist`
7. `ReplacementCapable`
8. `RoboticCapable`

The baseline workbook remains compatible before migration: if `STUDENTS` does not exist, the Python reader returns an empty student registry. Once the sheet exists, the headers are validated strictly.

Validated registry semantics:
- non-empty stable student ID
- unique student ID
- non-empty display name
- positive and unique student number
- placement end cannot be before placement start
- replacement capability defaults to `True` when blank
- robotic capability defaults to `False` when blank
- confirmed student capacity remains five daily timeslots and is not exposed as a new workbook field

The registration validator also checks the supervisor against the therapist registry when a supervisor is supplied. Duplicate student display names remain allowed because identity is kept separately.

The student registration preview workflow:
- creates a new `.xlsm` copy and never opens the baseline for writing
- creates the authoritative `STUDENTS` schema in the copy when it is not already present
- appends the validated student to the next logical row
- writes placement dates as real Excel date values; date display formatting is best-effort only so localized COM formatting cannot block safe data writeback
- uses the workbook's configured yes/no labels for replacement and robotic capability when possible
- verifies the complete student registry by reading the preview back
- hashes the source before and after the operation

## Important boundary

Preview writeback never changes the baseline workbook. Final menu/form UI remains deferred until these source/write contracts are stable.
