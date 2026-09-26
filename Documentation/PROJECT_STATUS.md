# Project Status

Last updated: 2026-09-26

## Current stage

The project is now in the outpatient scheduling integration phase.

Completed milestones:
- unified Python registration orchestration for patient, therapist and student previews
- registration menu/form contract in Python
- safe preview-only VBA registration menu installer
- central `frmRegistrationMenu` UserForm with buttons for patient, therapist, student and close
- patient registration UserForm (`frmNewPatient`) validated on the target Windows/Excel installation
- PR #14 merged into `main`
- full suite at PR #14 validation point: `262 passed, 3 warnings in 2.94s`

## Current branch / pull request stage

Branch: `feature/outpatient-scheduling`
PR #15: `Add outpatient scheduling domain support`
Status: draft / implementation in progress.

Latest confirmed full suite on this branch: `282 passed, 3 warnings in 2.93s`.

## Outpatient domain work completed on this branch

- Added `PatientType` with values `INPATIENT` and `OUTPATIENT`.
- Existing/legacy `Patient(...)` objects default to `INPATIENT` for backward compatibility.
- Added optional `hospital_mrn` to the patient model while keeping `patient_id` as the permanent internal identifier.
- Added `Patient.is_outpatient` convenience property.
- Domain rejects an outpatient marked infectious.
- Added explicit operational-semantics tests confirming that outpatient status does not remove a session from therapist workload, daily capacity, or replacement-candidate logic.
- Operational-semantics tests and full suite have been rerun successfully at the latest checkpoint above.

## Strict outpatient presentation rules

- An outpatient is visually **light blue only** in `THERAPIST DAILY`.
- An outpatient can never be infectious/yellow.
- An outpatient can never receive robotic treatment or a robotic extra programme.
- A recurring therapist/time slot may be shared with another patient on different weekdays.
- The colour shown in `THERAPIST DAILY` is determined by the patient actually active in that slot on the concrete date.
- Colours from another patient who shares the recurring slot on other weekdays must not leak into the outpatient day.
- If the same effective daily cell would contain both an inpatient and an outpatient at the same time, the renderer must reject the ambiguous cell-level colour instead of guessing.
- Active outpatient sessions create a blue daily target.
- Replaced outpatient sessions make the replacement destination blue.
- Cancelled/no-show outpatient sessions release the slot and do not claim an active blue target.

Implementation added:
- `Python/rehab_excel/outpatient_presentation.py`
- daily outpatient target calculation
- strict robotic-assignment validation for outpatients
- presentation-only `outpatient_light_blue` patch builder
- `Tests/test_outpatient_presentation.py`

## Recurring outpatient storage decision

`PATIENT_PLANNER` is an inpatient/hospitalized-patient planner and should not visibly contain outpatients.

Outpatients still need authoritative recurring source data. They cannot exist only in `THERAPIST DAILY` text because recurring materialization, workload, capacity, replacements and cancellations depend on persistent schedule rows.

Decision recorded in `Documentation/DECISIONS.md`:
- inpatient recurring rows remain in `PATIENT_PLANNER`
- outpatient recurring rows will use a separate authoritative source, provisionally `OUTPATIENT_SCHEDULE`
- the scheduling engine merges both into one `BaseScheduleEntry` stream
- patient type affects storage/view routing, not operational scheduling semantics

## Daily treatment cancellation work completed on this branch

A daily cancellation is a session-specific overlay. It does not modify the recurring/base schedule.

Added:
- `DailySessionCancellation`
- `SessionCancellationKind.DEPARTMENT_POSTPONED`
- `SessionCancellationKind.PATIENT_NO_SHOW`
- `DailySessionStatus.CANCELLED`
- `DailySessionState.releases_provider_slot`

Rules implemented:
- a cancellation targets one concrete session on one date
- another treatment for the same patient on the same day remains active unless separately cancelled
- cancellation has priority over a replacement overlay for the cancelled session
- cancelled/no-show treatment has no effective therapist/time for that session
- original recurring Session remains immutable
- the original therapist/time is released for use by another patient/replacement
- cancellation reason and kind remain available for future audit/statistics/rendering

## Outpatient rules still to integrate

Outpatients must:
- continue counting in time-share / productivity accounting
- persist safely through workbook reader/storage
- be excluded from inpatient-only / hospitalized-patient views
- be excluded from the inpatient patient-planner view
- be rendered through the actual `THERAPIST DAILY` preview/write workflow with light-blue fill

Important principle: outpatient status changes visibility/presentation/storage routing, not whether a session counts operationally.

Operational expectation: outpatients are about 15% max of the patient population, but this is not a hard validation ceiling.

## Next steps

1. Integrate outpatient blue presentation patches into the existing `THERAPIST DAILY` preview/write pipeline.
2. Ensure presentation-only Excel patches never overwrite cell values.
3. Add the `OUTPATIENT_SCHEDULE` read/merge contract without changing the current workbook yet.
4. Make reader/storage changes needed to persist patient type / hospital MRN safely.
5. Filter outpatients from inpatient-only views while preserving them in the unified scheduling stream.
6. Expose daily cancellation/no-show input in the operational Excel workflow.
7. Smoke-test the resulting preview workbook on the target Excel installation.

## Safety constraints

- Never write directly to the baseline `.xlsm` during development/smoke testing.
- Preview workflows must operate on copied `.xlsm` files only.
- Verify source workbook hash remains unchanged.
- Preserve and verify `xl/vbaProject.bin`.
- Keep existing scheduling behaviour backward-compatible unless an explicit business rule changes it.

## Recovery note

If project context is lost, start with:
1. this file
2. `Documentation/DECISIONS.md`
3. open pull requests
4. the latest merged PRs
5. the test suite

Do not infer undocumented business rules when one of these sources can confirm them.
