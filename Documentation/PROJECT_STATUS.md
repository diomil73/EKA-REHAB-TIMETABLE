# Project Status

Last updated: 2026-09-26

## Current stage

The project is in the registration UI integration phase.

Completed milestones:
- unified Python registration orchestration for patient, therapist and student previews
- registration menu/form contract in Python
- safe preview-only VBA registration menu installer
- central `frmRegistrationMenu` UserForm with buttons for patient, therapist, student and close
- real Excel smoke test confirmed the menu renders correctly on the target Windows/Excel installation
- patient registration UserForm (`frmNewPatient`) added behind the `Νέος ασθενής` menu action
- real Excel smoke test confirmed the patient form renders correctly, defaults to `Εσωτερικός`, can switch to `Εξωτερικός`, shows automatic locked PatientID, exposes optional hospital MRN, and disables inpatient-only fields for outpatients

## Current pull request

PR #14: `Add patient registration UserForm`
Branch: `feature/patient-registration-userform`
Status: Excel smoke test passed; ready for review/merge after tests are confirmed.

The current patient form is preview-only and deliberately performs no workbook write yet.

## Current patient-registration decisions

- Patient type is a combo box.
- Default patient type is `Εσωτερικός` to minimize clicks.
- User may switch to `Εξωτερικός` from the combo box.
- Outpatients are operationally expected to be a minority, about 15% max, but this is not a hard registration limit.
- `PatientID` should be generated automatically and remain a permanent internal system identifier.
- The hospital registry number (`ΑΜ Νοσοκομείου` / Hospital MRN) is a separate optional field.
- Hospital MRN may be filled in later without changing PatientID.
- Room and patient status choices come from `SETTINGS` columns H and F.
- Inpatient-only fields are disabled/cleared when patient type is switched to outpatient.

## Outpatient rules that still need implementation

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

## Next steps

1. Confirm focused and full pytest suites for PR #14.
2. Merge PR #14 once tests are green.
3. Start a separate outpatient scheduling PR touching the domain model, readers, schedule generation, workload/capacity/replacements and `THERAPIST DAILY` rendering.
4. After scheduling rules are stable, connect the patient UserForm to the existing safe Python registration backend.

## Safety constraints

- Never write directly to the baseline `.xlsm` during development/smoke testing.
- Preview workflows must operate on copied `.xlsm` files only.
- Verify source workbook hash remains unchanged.
- Preserve and verify `xl/vbaProject.bin`.
- Keep Excel/VBA compatibility with the target Office installation. The target setup rejected COM writes to some MSForms `Width`/`Height` properties, so final form sizing/styling is performed inside VBA `UserForm_Initialize`.

## Recovery note

If project context is lost, start with:
1. this file
2. `Documentation/DECISIONS.md`
3. open pull requests
4. the latest merged PRs
5. the test suite

Do not infer undocumented business rules when one of these sources can confirm them.
