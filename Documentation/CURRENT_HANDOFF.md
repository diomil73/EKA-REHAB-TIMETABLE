# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V81.xlsm`
- V73 scheduling/conflict behavior remains native green.
- V80 clean reset preview is native green.
- V81 MASTER toolbar/header pinning is native green.

## MASTER controls status

Native experiments V74-V79 are retired as sidebar/persistent-control attempts.

Final accepted MASTER UX:

- row 2 contains the four existing horizontal action buttons,
- row 4 contains the MASTER treatment/column headers,
- the top area is frozen through row 4,
- the application shell reapplies the pin whenever MASTER becomes active,
- no UserForm/sidebar/CommandBars mechanism is used.

Native V81 verification: green. Vertical scrolling keeps both the four MASTER buttons and row-4 treatment headers visible.

## Patient-program native verification

Confirmed green by the user:

1. edit-existing-program end to end,
2. existing program preload,
3. atomic resave without false self-conflict,
4. optional manual names persistence/reload for Ρομποτικό / Εργο / Λογο / Ψυχ.

## Retired single-treatment cleanup

Cleanup is complete and accepted.

Removed/retired source:

- `Python/rehab_excel/patient_scheduling_flow.py`
- `Python/rehab_excel/patient_scheduling_form_vba.py`
- obsolete `plan`, `validate`, `plan-vba`, `validate-vba`, `recurring-vba`, and `validate-recurring-vba` commands from `Python/tools/patient_scheduling_bridge.py`
- tests that exercised only the retired single-treatment flow/form

Retained current behavior:

- `auto-propose` / `auto-propose-vba` in `patient_scheduling_bridge.py`, used by the joint patient-program auto scheduler,
- current `PatientProgramRuntime` / `frmPatientScheduling`,
- legacy VBA component removal/verification in `patient_registration_scheduling_slice.py` so old workbooks cannot retain `PatientSchedulingRuntime`.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- cleanup/reset focused gate: `42 passed in 3.70s`
- row-4 toolbar/header pinning focused source gate: `43 passed in 1.92s`
- V81 native toolbar/header pinning gate: green
- optional manual names persistence/reload: native green
- full regression before cleanup: `722 passed, 19 warnings`
- **full regression after cleanup: `699 passed, 19 warnings`**

The drop from 722 to 699 is expected because obsolete tests for the removed single-treatment code were deleted. No post-cleanup regression was reported.

## Cancellation slot release status

The first outpatient roadmap candidate is already implemented and regression-covered in source:

- cancellations are passed into therapist/patient/student availability checks,
- cancelled occurrences are excluded from therapist/student operational workload,
- replacement candidate capacity reflects the released timeslot,
- `Tests/test_cancellation_releases_replacement_capacity.py` verifies both workload release and exact-slot reuse.

No duplicate implementation is needed here. Native operational smoke remains useful later, but the source gap is closed.

## Current focus

**Outpatient recurring schedule create/edit end to end against `OUTPATIENT_SCHEDULE`.**

Immediate work:

1. inspect the existing `frmOutpatientSchedule` create flow and bridge,
2. verify whether user-facing edit/preload already exists,
3. if missing, add edit/preload while reusing `OutpatientScheduleRequest.target_base_entry_id`,
4. focused source tests,
5. build a native smoke preview and verify create + edit on Windows/Excel.

Required invariants:

- outpatient recurring rows stay in `OUTPATIENT_SCHEDULE`, not `PATIENT_PLANNER`,
- robotic treatment remains unavailable to outpatients,
- conflict/suggestion safeguards remain intact,
- preview workflows never mutate the source workbook,
- VBA is preserved in generated previews.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative for inpatient recurring schedules; `OUTPATIENT_SCHEDULE` authoritative for outpatient recurring schedules.
- Conflict override must be explicit; return-for-correction performs zero writes.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
