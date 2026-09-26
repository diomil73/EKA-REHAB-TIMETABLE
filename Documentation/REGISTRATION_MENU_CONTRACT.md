# Registration menu/form contract

This contract is the thin UI-facing layer between a future central Excel/Windows menu and the proven registration orchestration backend.

## Central registration choices

The stable menu actions are:

1. `new_patient` — Νέος ασθενής
2. `new_therapist` — Νέος θεραπευτής
3. `new_student` — Νέος φοιτητής

`rehab_excel.registration_menu` exposes these actions as `RegistrationFormSpec` objects. A UI can use the specs to render the correct fields without embedding registration business rules in VBA or another presentation layer.

## Patient form

Fields:
- `Patient ID` — required text
- `Ονοματεπώνυμο` — required text
- `Θάλαμος` — optional, backed by the configured `rooms` source
- `Λοιμώδης ασθενής` — boolean, defaults to false
- `Κατάσταση` — optional, backed by the configured `patient_statuses` source

## Therapist form

Fields:
- `Ονοματεπώνυμο` — required text

Robotic capability is deliberately not exposed by the form yet. The current workbook has no confirmed authoritative therapist-capability storage location, so the existing safe backend rejects attempts to persist that capability rather than silently dropping it.

## Student form

Fields:
- `Student ID` — required text
- `Ονοματεπώνυμο` — required text
- `Αριθμός φοιτητή` — required integer
- `Έναρξη πρακτικής` — required date
- `Λήξη πρακτικής` — required date
- `Επόπτης θεραπευτής` — optional, backed by the registered `therapists` source
- `Δυνατότητα αναπληρώσεων` — boolean, defaults to true
- `Ρομποτική αποκατάσταση` — boolean, defaults to false

## Parsing boundary

`build_registration_request()` converts raw UI values into the existing domain request classes:
- `NewPatientRequest`
- `NewTherapistRequest`
- `NewStudentRequest`

This layer only handles presentation-shape parsing such as trimming text, parsing integer/date values and accepting normal boolean representations. It intentionally does not duplicate authoritative registration validation.

Dates accepted at the form boundary:
- `DD/MM/YYYY`
- `YYYY-MM-DD`
- native Python `date` / `datetime` values

Boolean parsing accepts booleans, 0/1 and common yes/no strings including Greek `ναι`, `όχι` and unaccented `οχι`.

## Safety boundary

After a form produces a domain request, the caller must pass that request to `create_registration_preview()`.

The existing orchestrator remains responsible for routing to the proven patient/therapist/student preview workflow. The underlying workflows remain responsible for authoritative validation, native Excel COM writeback to a copy, source hashing and read-back verification.

This contract does not yet add a VBA UserForm to the binary `.xlsm`, and it does not permit direct writes to the baseline workbook.
