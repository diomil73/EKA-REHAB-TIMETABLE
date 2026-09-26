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
Status: implementation in progress; draft PR to be opened.

## Outpatient domain work completed on this branch

- Added `PatientType` with values `INPATIENT` and `OUTPATIENT`.
- Existing/legacy `Patient(...)` objects default to `INPATIENT` for backward compatibility.
- Added optional `hospital_mrn` to the patient model while keeping `patient_id` as the permanent internal identifier.
- Added `Patient.is_outpatient` convenience property.

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

## Outpatient rules still to implement

Outpatients must:
- participate normally in scheduling
- count in therapist workload
- count in time-share / productivity accounting
- count in provider capacity checks
- participate in replacement logic exactly like inpatients
- follow the same recurring days and treatment times
- appear in `THERAPIST DAILY`
- render with a light-blue cell color in `THERAPIST DAILY`

Outpatients must not:
- appear in inpatient-only / hospitalized-patient views
- appear in the inpatient patient-planner view

Important principle: outpatient status changes visibility/presentation, not whether a session counts operationally.

Operational expectation: outpatients are about 15% max of the patient population, but this is not a hard validation ceiling.

## Next steps

1. Run focused tests for patient type and daily session-state changes.
2. Run the full pytest suite and fix any compatibility regressions.
3. Make reader/storage changes needed to persist patient type / hospital MRN safely.
4. Ensure outpatient sessions participate unchanged in workload, capacity and replacement calculations.
5. Filter outpatients only from inpatient-only patient/planner views.
6. Render outpatient sessions in `THERAPIST DAILY` with light-blue fill.
7. Expose daily cancellation/no-show input in the operational Excel workflow.
8. Smoke-test the resulting preview workbook on the target Excel installation.

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
