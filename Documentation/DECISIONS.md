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

### Recurring schedule storage
`PATIENT_PLANNER` remains the inpatient/hospitalized-patient planner and must not become the visible home of outpatient schedules.

However, outpatient recurring appointments still need an authoritative persistent source. They must not exist only as text in `THERAPIST DAILY`, because workload, capacity, replacements, cancellations and future-day materialization all depend on the recurring source data.

Architecture decision:
- keep inpatient recurring rows in `PATIENT_PLANNER`
- store outpatient recurring rows in a separate authoritative source, provisionally named `OUTPATIENT_SCHEDULE`
- the scheduling engine reads/merges both sources into the same `BaseScheduleEntry` stream
- after merging, inpatient and outpatient sessions use the same operational scheduling engine
- `THERAPIST DAILY` remains the only normal operational view where outpatient names are shown

The outpatient source may later be hidden/protected in Excel, but it remains data storage, not a second scheduling engine.

### Strict visual / clinical rule
An outpatient is **light blue only** in `THERAPIST DAILY`.

An outpatient:
- is never infectious/yellow
- never receives robotic treatment / robotic extra programme
- must not inherit yellow, pink, or any other patient colour from another patient who shares the same recurring therapist/time slot on other weekdays

When a therapist/time slot is shared by different patients on different weekdays, `THERAPIST DAILY` presentation is determined by the patient actually active on the concrete date.

If the same effective daily cell would contain both an inpatient and an outpatient simultaneously, cell-level colouring is ambiguous and the renderer must reject the situation rather than guess a colour.

### Daily postponement / no-show rule
The recurring outpatient appointment remains in the base schedule, but any specific occurrence may be cancelled for that date without deleting or moving the recurring slot.

A daily cancelled occurrence must:
- be marked with a reason, distinguishing at least `department_cancelled` from `patient_no_show`
- stop counting as delivered treatment for that date
- release the therapist timeslot so it can be used for replacement work if needed
- preserve the original recurring appointment for future dates
- remain auditable in daily history/statistics

This is a daily operational override, not a change to the patient's base programme.

## Global replacement optimization

### Decision
When more than one therapist absence creates displaced sessions on the same day, the system must not solve replacements greedily one therapist or one patient at a time. All known therapist absences for the day must be entered first, then all displaced/orphan sessions must be evaluated as one combined assignment problem before authoritative replacements are committed.

### Objective order
The optimizer should prefer, in order:

1. place as many displaced patients as possible and minimize unresolved orphans,
2. preserve each patient's original treatment time whenever feasible,
3. protect scarce placements so a flexible patient does not consume the only viable slot of a constrained patient,
4. minimize unnecessary patient moves and schedule disruption,
5. respect workload, provider capability, student limits, existing daily facts and all hard clinical/operational constraints,
6. use softer preferences and clinical convenience rules only after hard feasibility is satisfied.

### Scarcity principle
Candidate slots/providers must be evaluated globally. A placement that is locally valid can still be globally harmful if it removes the only feasible option for another displaced patient. The engine therefore needs awareness of candidate scarcity, not just candidate score for the current patient.

### Clinical knowledge boundary
Some placement rules depend on patient-specific clinical or operational knowledge that may not yet exist as structured data, such as desirable adjacency to another therapy or use of special equipment. These rules must be modelled explicitly where known as hard or soft constraints/preferences. Where knowledge is not encoded, the system must preserve manual user control rather than pretend to optimize what it cannot know.

### Manual override
Manual changes, accepted over-capacity placements and accepted same-time double treatment remain allowed only with explicit warning and final user confirmation. Automatic optimization should produce the best feasible proposal, not remove human authority.

### Architectural direction
The intended backend is a global constraint/assignment optimizer over all displaced sessions for the concrete date, not a sequence of independent replacement calls. A practical implementation may use weighted bipartite matching, min-cost flow, or a constraint solver, provided the locked business priorities above are testable and deterministic.

## Registration UI

- Central UserForm: `frmRegistrationMenu`
- Patient UserForm: `frmNewPatient`
- Current patient form is preview-only and must not write to the workbook yet.
- Room options come from `SETTINGS` column H.
- Patient status options come from `SETTINGS` column F.
- Inpatient-only controls are disabled/cleared when type changes to outpatient.

## Daily operational UI terminology

User-visible daily operational fields must use the agreed Greek business terminology only. Internal technical identifiers, source labels, command names or English implementation strings must never leak into normal Excel UI fields such as `Αιτία`.

For replacement rows created by the `DAILY_INPUT` therapist-absence workflow, legacy technical values such as `DAILY_INPUT therapist absence` or `temporary replacement` are rendered to the user as `Απουσία θεραπευτή`. Internal immutable IDs may remain technical because they are not user-facing labels.

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

Documentation is part of the implementation, not a cleanup task after implementation.

Whenever any meaningful development occurs, update the appropriate project records in the same work cycle. This includes:

- implemented behaviour or changed semantics,
- new tests or regression findings,
- compatibility workarounds,
- current validation state and native smoke results,
- defects, risks and technical debt,
- newly discovered necessities,
- ideas or requirements that we decide should later be built,
- functionality that should be removed, retired or replaced,
- fixes that are known but intentionally deferred,
- changes to implementation order or project priorities.

Use the files by purpose:

- `Documentation/DECISIONS.md` for durable business/architecture rules and agreed constraints,
- `Documentation/PROJECT_STATUS.md` for current progress, open risks, validation state and immediate roadmap,
- `Documentation/CURRENT_HANDOFF.md` for the exact resume point, latest checkpoints and commands/scenarios needed to continue safely,
- `Documentation/IMPLEMENTATION_MAP.md` when ownership/location of implemented capability changes,
- `CHANGELOG.md` for user-visible or implementation-relevant changes that belong in release history.

Do not rely on chat history as the only record of a decision, requirement, idea, removal, deferred fix or newly discovered need. If it may affect future implementation, testing or recovery after context loss, it must be written into the repository documentation.