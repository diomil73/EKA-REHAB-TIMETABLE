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
- obsolete single-treatment bridge commands and tests.

Retained current patient-program behavior and legacy component cleanup.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- cleanup/reset focused gate: `42 passed in 3.70s`
- row-4 toolbar/header pinning focused source gate: `43 passed in 1.92s`
- V81 native toolbar/header pinning gate: green
- optional manual names persistence/reload: native green
- full regression before cleanup: `722 passed, 19 warnings`
- **full regression after cleanup: `699 passed, 19 warnings`**

## Cancellation slot release status

Cancellation propagation is already implemented and regression-covered in source:

- cancellations are passed into therapist/patient/student availability checks,
- cancelled occurrences are excluded from therapist/student operational workload,
- replacement candidate capacity reflects the released timeslot,
- `Tests/test_cancellation_releases_replacement_capacity.py` verifies workload release and exact-slot reuse.

## Locked daily-operations rule: cancellations + therapist absences

Patient cancellations for both inpatients and outpatients are operationally critical because they release same-day timeslots that can be reused for:

- patients displaced by an absent therapist,
- redistribution of active patients across therapists,
- new patients when capacity remains.

The workflow must NOT depend on one fixed chronological entry order. Therapist absences are often known earlier, while patient cancellations may arrive much later. Therefore both event types must be accepted at any time and the daily plan must be recomputable/idempotent from the latest facts.

Required evaluation model whenever the operational plan is refreshed:

1. start from the recurring/base schedule for the selected date,
2. apply all known patient cancellations/no-shows for that date and release their occupied slots,
3. apply all known therapist absences and mark those providers unavailable,
4. identify the remaining displaced/unscheduled patients,
5. redistribute those existing affected patients into the net free capacity,
6. only after that expose remaining capacity for new-patient assignment.

Important temporal behavior:

- if therapist absence is entered first, the system may create a preliminary redistribution,
- if a patient cancellation arrives later, refreshing/recalculating the same date must immediately release the slot and allow a better redistribution,
- if a cancellation is removed/corrected, recalculation must reclaim the slot and revalidate assignments,
- base recurring schedules must not be destructively edited to represent a one-day cancellation or absence.

This is now a core invariant for the future daily allocation engine and native DAILY_INPUT/absence workflows.

## Current development focus

**Outpatient recurring schedule create/edit end to end against `OUTPATIENT_SCHEDULE`.**

The backend already supports update by `target_base_entry_id`; user-facing edit/preload support is being added by reusing the same form/backend rather than introducing a second writer.

After the outpatient create/edit gate, the next operational milestone is the dynamic daily recomputation flow above, with native verification that cancellations and therapist absences can be entered in either order.

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
- One-day cancellations and therapist absences are overlays on the daily operational state, not destructive edits of recurring source schedules.
- Existing displaced patients have allocation priority over newly added patients when released capacity is redistributed.
- Conflict override must be explicit; return-for-correction performs zero writes.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
