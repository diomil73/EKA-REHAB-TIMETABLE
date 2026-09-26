# Project Decisions

This file records business and architecture decisions that should survive loss of chat/session context.

## Patient identity

### Decision
Use two separate identifiers:

- `PatientID`: permanent internal system identifier, generated automatically.
- `HospitalMRN` / `ΑΜ Νοσοκομείου`: optional hospital registry number, entered when known and editable later.

### Rationale
`PatientID` is used as the stable key across schedules, absences, replacements, history and future related records. It must not be replaced later by another identifier. Hospital MRN is external business data and may be unknown at first.

## Patient type

### Values
- `Εσωτερικός`
- `Εξωτερικός`

### UI rule
The patient registration combo box defaults to `Εσωτερικός` to minimize routine clicks. The user can switch to `Εξωτερικός` from the drop-down.

### Operational ratio
Outpatients are expected to be about 15% max of the total patient population. This is a practical operating ratio, not a hard validation ceiling.

## Outpatient scheduling semantics

### Core rule
An outpatient is fully real for scheduling and accounting. Patient type affects visibility and presentation, not operational counting.

### Must count normally in
- therapist daily workload
- provider capacity
- time-share / productivity accounting
- recurring treatment schedules
- replacement calculations
- daily therapist schedules

### Visibility
Outpatients should not appear in inpatient-only patient views or the inpatient patient planner.

Outpatients should appear in `THERAPIST DAILY`.

### Visual rule
Outpatient cells in `THERAPIST DAILY` must use a light-blue fill so they are visually distinguishable from inpatients.

## Registration UI

- Central UserForm: `frmRegistrationMenu`
- Patient UserForm: `frmNewPatient`
- Current patient form is preview-only and must not write to the workbook yet.
- Room options come from `SETTINGS` column H.
- Patient status options come from `SETTINGS` column F.
- Inpatient-only controls are disabled/cleared when type changes to outpatient.

## Excel/VBA compatibility

The target Excel/pywin32 installation rejected programmatic COM writes to some MSForms `Width`/`Height` properties.

Decision:
- use COM only to create the form and controls
- perform final sizing/styling from VBA `UserForm_Initialize`

This pattern is now the preferred approach for future dynamically installed VBA UserForms unless a later compatibility test proves a better method.

## Development safety

- Never modify the baseline `.xlsm` directly during development.
- Create preview copies first.
- Verify source hash remains unchanged.
- Preserve VBA project content.
- Keep entity-specific backend validation authoritative. UI validation is convenience/presentation only.

## Documentation maintenance rule

Whenever a major business rule, architectural choice, compatibility workaround, or project-stage transition is agreed:
- update `Documentation/DECISIONS.md` if it is a durable decision
- update `Documentation/PROJECT_STATUS.md` if it changes current progress, next steps, open risks or validation state

These files are the recovery anchor for continuing the project after loss of conversational context.
