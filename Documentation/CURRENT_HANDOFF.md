# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `Excel/previews/DAILY_CHANGE_LEDGER_V84.xlsm`
- V73 scheduling/conflict behavior remains native green.
- V80 clean reset preview is native green.
- V81 MASTER toolbar/header pinning is native green.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- cleanup/reset focused gate: `42 passed in 3.70s`
- row-4 toolbar/header pinning focused source gate: `43 passed in 1.92s`
- full regression after cleanup: `699 passed, 19 warnings`
- reversible daily change ledger focused source gate: `10 passed in 3.26s`
- replacement load + mixed outpatient/inpatient regression gate: `31 passed in 2.04s`
- MASTER daily overlay + ledger restore source gate: **24 passed, 5 warnings in 4.39s**
- V82 guided replacement native build: green
- V82 individual ledger restore: **native green**
- V83 MASTER active-overlay sync + restore: **native green**
- V84 time-axis display + restore: **native green**

The 5 warnings are `PytestUnraisableExceptionWarning` messages from Python 3.14 `ZipFile.__del__` during workbook test teardown. They did not fail the gate and are not currently treated as a functional blocker; keep them on the technical-cleanup list.

## V84 native guided replacement checkpoint

Stress case:

- date: `2026-10-19`
- patient: `ΜΕΝΤΗΣ`
- original assignment: `Σκαρώνης 09:15`
- selected replacement: `Μηλιδάκης 10:45`
- source workbook unchanged,
- VBA preserved,
- ledger installed.

Native behavior confirmed:

- `THERAPIST_DAILY` shows the moved patient at `Μηλιδάκης 10:45` while keeping the useful source-position trace,
- `MASTER` shows the same effective daily assignment while the overlay is active,
- individual `Επαναφορά` returns both `MASTER` and `THERAPIST_DAILY` to `Σκαρώνης 09:15`,
- left-side timetable times are now readable as `hh:mm`; the `####` issue is closed.

Open native cleanup still visible in V84:

- ΜΕΝΤΗΣ appears orange at the destination even though he is not robotic,
- the obsolete `ΑΝΑΝΕΩΣΗ` button is still present.

## Current styling + legacy-button source fix

Root cause of the orange leak:

- rewritten destination cells could inherit a stale whole-cell orange font from an existing robotic presentation,
- a new non-robotic line with no explicit text-run color then inherited that orange.

Source fix added:

- native Excel rich-text writes now reset the rewritten cell to default black when no whole-cell font role is requested,
- semantic rich-text runs are then reapplied, so only actual robotic lines receive `robotic_orange`,
- this preserves mixed-cell per-line semantics instead of leaking old whole-cell formatting.

Legacy cleanup added:

- ledger preview installation now removes only shapes on `THERAPIST_DAILY` whose exact caption is `ΑΝΑΝΕΩΣΗ`,
- ledger controls are not targeted.

Commits:

- `f4e382e0f1039cbd9582e9138a7dacde5eabd404` rich-text stale font reset
- `d83b36d0d38d782df9e3cc094eb0da18498906a1` remove legacy refresh button
- `1785786c05f95ac39aa8f2e17d9efb6e8e46325c` refresh-button regression
- `9fc4e948d0065dfbad838730584d2832ef8adef7` rich-text reset regression

Status: **implemented, focused/native verification pending**.

## Locked orphan workflow

Do not forget the daily list of unresolved/displaced patients (`Ορφανοί`). Every orphan for the concrete date must end in one of two operational outcomes:

- **placed** into a valid therapist/time slot, or
- **removed for that date only** with reason `Αναστολή θεραπείας`.

Rules:

- this is a daily overlay/status, never a destructive edit of the recurring/base schedule,
- an orphan must not silently disappear from the operational plan,
- `Αναστολή θεραπείας` must feed daily statistics as a treatment not performed for that explicit reason,
- orphan resolution participates in the same recomputable daily-state engine,
- displaced existing patients have priority over new-patient allocation.

## Locked daily-operations rule: cancellations + therapist absences

Patient cancellations for both inpatients and outpatients release same-day timeslots. Daily computation must be independent of entry order:

1. recurring/base schedule for selected date,
2. apply cancellations/no-shows and release slots,
3. apply therapist absences,
4. identify displaced/unscheduled patients,
5. redistribute existing affected patients into net free capacity,
6. resolve any remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Recomputation must be idempotent and recurring schedules must never be destructively edited for one-day facts.

## Locked replacement load semantics

Operational therapist load is occupied daily **timeslots**, not patient/session count.

- `active_timeslots` is authoritative load/capacity,
- multiple patients sharing one clock slot count as multiple sessions but one occupied timeslot,
- replacement ranking sorts primarily by occupied timeslots,
- candidate display distinguishes `load x/y` from `patients n`.

## Mixed inpatient/outpatient presentation rule

A mixed inpatient/outpatient effective slot must not block a valid operational preview. Do not apply one whole-cell outpatient fill to a mixed cell; keep it neutral until per-patient styling exists.

## Reversible daily change ledger

Required section in `THERAPIST_DAILY`:

`Αλλαγές για την dd/mm/yy`

Each move preserves patient, original/effective therapist and time, reason/source, stable id and status.

Controls:

- per-row `Επαναφορά`,
- selection,
- `Επαναφορά επιλεγμένων`,
- `Επαναφορά όλων`.

Production restore semantics:

- remove/reverse the date-specific override and recompute the day,
- never reconstruct or rewrite recurring schedule from UI text,
- idempotent,
- stale undo blocked,
- recalculate capacity/conflicts/replacements after restore,
- bulk restore validated and safe.

Prototype currently uses preview snapshots plus a MASTER recurring refresh hook; after UX acceptance it must be connected to authoritative overlay removal + full daily recomputation.

## Current development focus / exact resume point

Priority order:

1. **Run focused regressions for stale-font reset + legacy `ΑΝΑΝΕΩΣΗ` removal.**
2. Build fresh preview and verify ΜΕΝΤΗΣ is no longer orange at Μηλιδάκης 10:45.
3. Verify `ΑΝΑΝΕΩΣΗ` is gone and ledger controls remain.
4. Verify individual restore still works and MASTER/THERAPIST_DAILY remain synchronized.
5. Implement/verify orphan queue semantics, including `Αναστολή θεραπείας` and statistics.
6. Test selected/all restore with multiple changes.
7. Replace preview snapshot restore with authoritative overlay removal + daily recompute.
8. Resume outpatient recurring create/edit native gate.
9. Clean up Python 3.14 ZipFile teardown warnings when it does not distract from functional milestones.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Outpatients cannot receive Ρομποτικό.
- Nonblank manual constraints are never silently overwritten.
- Standard duration 45 minutes; Ανακλινόμενο 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative for inpatient recurring schedules; `OUTPATIENT_SCHEDULE` authoritative for outpatient recurring schedules.
- one-day cancellations, absences, replacements, suspensions and redistributions are daily overlays, not destructive recurring edits,
- operational load/capacity is timeslot-based,
- existing displaced patients precede new patients,
- unresolved orphans must be placed or explicitly suspended for the date,
- `MASTER` and `THERAPIST_DAILY` must agree on the effective daily state.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
