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

Latest confirmed full suite on this branch: `310 passed, 3 warnings in 3.69s`.

The three warnings remain the known openpyxl/zipfile warnings already seen in prior checkpoints; no new test regression is represented by them.

## Outpatient domain work completed on this branch

- Added `PatientType` with values `INPATIENT` and `OUTPATIENT`.
- Existing/legacy `Patient(...)` objects default to `INPATIENT` for backward compatibility.
- Added optional `hospital_mrn` to the patient model while keeping `patient_id` as the permanent internal identifier.
- Added `Patient.is_outpatient` convenience property.
- Domain rejects an outpatient marked infectious.
- Added explicit operational-semantics tests confirming that outpatient status does not remove a session from therapist workload, daily capacity, or replacement-candidate logic.

## Strict outpatient presentation rules

- An outpatient is visually **light blue only** in `THERAPIST DAILY`.
- An outpatient can never be infectious/yellow.
- An outpatient can never receive robotic treatment or a robotic extra programme.
- A recurring therapist/time slot may be shared with another patient on different weekdays.
- The colour shown in `THERAPIST DAILY` is determined by the patient actually active in that slot on the concrete date.
- Colours from another patient who shares the recurring slot on other weekdays must not leak into the outpatient day.
- If the same effective daily cell would contain both an inpatient and an outpatient simultaneously, the renderer rejects the ambiguous cell instead of guessing.
- Active outpatient sessions create a blue daily target.
- Replaced outpatient sessions make the replacement destination blue.
- Cancelled/no-show outpatient sessions release the slot and do not claim an active blue target.
- Stale outpatient blue is cleared only when the existing fill is exactly the outpatient light-blue role; yellow, pink and unrelated fills are left untouched.

Implementation added:
- `Python/rehab_excel/outpatient_presentation.py`
- daily outpatient target calculation
- strict robotic-assignment validation for outpatients
- presentation-only `outpatient_light_blue` patch builder
- stale-blue cleanup patches using `clear_fill`
- `Tests/test_outpatient_presentation.py`

## Daily preview integration completed on this branch

- `WriteIntent.PRESENTATION` in the native Excel backend is formatting-only and never writes `cell.Value`.
- Added semantic `clear_fill` support for safe stale outpatient-blue cleanup.
- Added `Python/rehab_excel/daily_preview_composer.py`.
- Presentation patches are merged into existing cell value patches instead of creating conflicting duplicate writes.
- If a replacement/cancellation patch already changes cell text, outpatient blue is folded into that same patch.
- Added dedicated tests in `Tests/test_presentation_write_intent.py` and `Tests/test_daily_preview_composer.py`.

## Recurring outpatient storage contract

`PATIENT_PLANNER` remains an inpatient/hospitalized-patient planner and should not visibly contain outpatients.

Outpatients still need authoritative recurring source data. They cannot exist only in `THERAPIST DAILY` text because recurring materialization, workload, capacity, replacements and cancellations depend on persistent schedule rows.

Decision recorded in `Documentation/DECISIONS.md`:
- inpatient recurring rows remain in `PATIENT_PLANNER`
- outpatient recurring rows use a separate authoritative source named `OUTPATIENT_SCHEDULE`
- `OUTPATIENT_SCHEDULE` is one row per recurring treatment with columns `PatientID`, `Ασθενής`, `Θεραπεία`, `Ώρα`, `Ημέρες`, `Θεραπευτής`
- absence of `OUTPATIENT_SCHEDULE` is backward-compatible and returns no outpatient rows
- the scheduling layer merges inpatient and outpatient rows into one `BaseScheduleEntry` stream
- patient type affects storage/view routing, not operational scheduling semantics

Implementation added:
- `Python/rehab_excel/outpatient_schedule_source.py`
- `read_outpatient_schedule()`
- `read_unified_base_schedule()`
- strict identity, day-pattern, outpatient-only and no-robotic validation
- `Tests/test_outpatient_schedule_source.py`

## Patient registry storage extension

Added `Python/rehab_excel/patient_registry_source.py` as the backward-compatible patient registry reader for the outpatient pipeline.

Rules:
- legacy five-column `PATIENTS` remains valid
- legacy rows default to `PatientType.INPATIENT`
- optional patient type headers accepted: `PatientType`, `ΤύποςΑσθενή`, `Τύπος Ασθενή`
- optional hospital MRN headers accepted: `HospitalMRN`, `ΑΜ Νοσοκομείου`, `ΑΜΝοσοκομείου`
- Greek patient-type labels are Unicode-normalized so accented/case/final-sigma variants are accepted safely
- unknown patient-type text raises instead of silently defaulting to inpatient
- outpatient + infectious remains impossible through domain validation

Added `Tests/test_patient_registry_source.py`.

## Real operational caller migration

`Python/tools/apply_daily_input.py`, `Python/tools/apply_daily_input_replacements.py`, and `Python/tools/apply_replacement_choice.py` use the unified outpatient flow:
- patients via `read_patient_registry()`
- recurring programme via `read_unified_base_schedule()`
- outpatient blue composition through `compose_outpatient_daily_plan()` before native Excel writeback
- clean safety handling for outpatient registry/schedule/presentation/composition errors

Added integration wiring tests:
- `Tests/test_apply_daily_input_outpatient_integration.py`
- `Tests/test_replacement_tools_outpatient_integration.py`

Existing workbooks without the new `OUTPATIENT_SCHEDULE` sheet continue to behave as before because the outpatient source is optional.

## Patient registration preview schema integration

`NewPatientRequest` carries optional `patient_type` and `hospital_mrn` while defaulting to `INPATIENT` for legacy callers.

`Python/rehab_excel/patient_registration.py` prepares outpatient-compatible storage only on the copied preview workbook:
- locates or creates PATIENTS extension columns for `PatientType` and `HospitalMRN`
- creates `OUTPATIENT_SCHEDULE` if missing, with the authoritative six-column header contract
- writes `Εσωτερικός` / `Εξωτερικός` and optional hospital MRN
- clears inpatient-only room/infectious/status fields for outpatient registrations
- verifies the new patient through `read_patient_registry()` including patient type and MRN
- continues hashing the source before/after so the baseline remains unchanged

Registration validation rejects outpatient + infectious and outpatient + inpatient-room combinations before Excel writeback.

Added `Tests/test_patient_registration_outpatient_contract.py`.

## Inpatient-only visibility boundary

No additional filter was added to `THERAPIST_DAILY` or `DAILY_INPUT` because those are operational views and must include outpatients.

The inpatient-only boundary is structural:
- `PATIENT_PLANNER` remains the hospitalized/inpatient recurring source
- outpatient recurring rows live only in `OUTPATIENT_SCHEDULE`
- operational callers merge both sources only where scheduling/workload/replacement logic requires them

This avoids accidentally removing outpatients from operational calculations while keeping them out of the hospitalized planner.

## Daily treatment cancellation work

A daily cancellation is a session-specific overlay. It does not modify the recurring/base schedule.

Core domain already includes:
- `DailySessionCancellation`
- `SessionCancellationKind.DEPARTMENT_POSTPONED`
- `SessionCancellationKind.PATIENT_NO_SHOW`
- `DailySessionStatus.CANCELLED`
- `DailySessionState.releases_provider_slot`

Operational input integration now added, pending test confirmation:
- `DAILY_INPUT` patient status `ΑΝΑΒΟΛΗ ΤΜΗΜΑΤΟΣ` maps to `DEPARTMENT_POSTPONED`
- `DAILY_INPUT` patient status `ΔΕΝ ΠΡΟΣΗΛΘΕ` maps to `PATIENT_NO_SHOW`
- both statuses require one concrete time and must resolve to exactly one scheduled session for that patient/date/time
- cancellation rows cannot use `Όλη ημέρα`
- cancellation has priority over therapist absence in the replacement queue, so a cancelled/no-show session is not sent for therapist replacement
- `apply_daily_input.py` and `apply_daily_input_replacements.py` now pass cancellation overlays into daily state construction

Added `Tests/test_daily_input_cancellations.py`.

Rules preserved:
- another treatment for the same patient on the same day remains active unless separately cancelled
- cancellation has priority over a replacement overlay for the cancelled session
- cancelled/no-show treatment has no effective therapist/time for that session
- original recurring Session remains immutable
- original therapist/time is released operationally
- cancellation reason and kind remain available for audit/statistics/rendering

## Outpatient rules still to integrate

Remaining work includes:
- ensure released cancellation slots are reflected in replacement capacity/availability calculations, not only daily state/queue priority
- add the actual outpatient recurring schedule entry/update workflow for `OUTPATIENT_SCHEDULE`
- smoke-test outpatient registration and operational previews on the target Excel installation

Important principle: outpatient status changes visibility/presentation/storage routing, not whether a session counts operationally.

Operational expectation: outpatients are about 15% max of the patient population, but this is not a hard validation ceiling.

## Next steps

1. Run `Tests/test_daily_input_cancellations.py` plus related daily-input/queue tests and the full suite.
2. Integrate explicit cancellation slot release into replacement option capacity/availability calculations.
3. Add the actual outpatient recurring schedule entry/update workflow for `OUTPATIENT_SCHEDULE`.
4. Smoke-test an outpatient registration preview on the target Excel installation and verify source hash/VBA preservation/new columns/new sheet.
5. Smoke-test a daily cancellation/no-show operational preview on the target Excel installation.

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
