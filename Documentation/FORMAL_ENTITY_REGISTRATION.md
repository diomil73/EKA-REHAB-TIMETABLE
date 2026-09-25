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

Preview writeback never changes the baseline workbook. Student Excel storage is still deliberately undefined until an authoritative registry location is agreed explicitly. Final menu/form UI is also deferred until these source/write contracts are stable.
