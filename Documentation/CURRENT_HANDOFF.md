# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest native preview tested: `Excel/previews/DAILY_CHANGE_LEDGER_V85.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement load + mixed outpatient/inpatient regression: `31 passed in 2.04s`
- MASTER daily overlay + ledger restore source gate: `24 passed, 5 warnings in 4.39s`
- V82 individual ledger restore: native green
- V83 MASTER active-overlay sync + restore: native green
- V84 time-axis display + restore: native green
- V85 stale-orange cleanup + restore: native green

Python 3.14 `ZipFile.__del__` warnings remain non-blocking technical debt.

## Native stress case now green

Case:

- date `2026-10-19`
- patient `ΜΕΝΤΗΣ`
- original `Σκαρώνης 09:15`
- replacement `Μηλιδάκης 10:45`

Confirmed native:

- `THERAPIST_DAILY` and `MASTER` agree on the effective daily assignment,
- source-position trace remains visible and is accepted UX,
- individual `Επαναφορά` returns both views to the original assignment,
- left-side time axis shows readable `hh:mm` with no `####`,
- moved non-robotic ΜΕΝΤΗΣ no longer inherits stale robotic-orange styling.

## Legacy `ΑΝΑΝΕΩΣΗ` button: deferred/manual cleanup

The obsolete Form Control still remains despite targeted source cleanup attempts.

Native diagnostic identified it as:

- Form Control / Shape type `8`,
- name `Button 1`,
- caption `ΑΝΑΝΕΩΣΗ`,
- macro `...!RefreshAll`.

Decision: **do not spend more engineering time on this now**. It is cosmetic, not a functional blocker. The user will remove it manually in Excel later if needed. Do not reopen this item unless it becomes operationally relevant.

## ACTIVE development focus: orphan workflow + statistics

The next substantive milestone is the daily `Ορφανοί` queue.

Locked rule: every orphan for the concrete date must resolve as exactly one of:

- placed into a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Required semantics:

- no orphan silently disappears,
- `Αναστολή θεραπείας` is a date-specific daily fact/overlay, never a destructive recurring edit,
- suspension must be auditable,
- suspension must feed daily statistics as treatment not performed with explicit reason,
- displaced existing patients have priority over new-patient allocation,
- recomputation must remain idempotent and order-independent.

## Locked daily recomputation order

1. recurring/base schedule for concrete date,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

## Reversible daily change ledger

Native green:

- active move rendering,
- MASTER synchronization,
- source trace,
- per-row restore,
- time-axis display,
- stale-orange cleanup.

Still open:

- selected/all restore with multiple simultaneous changes,
- production restore via authoritative overlay removal + full daily recompute instead of snapshot restoration.

## Exact resume point

Priority order:

1. Inspect existing daily-state/cancellation/statistics architecture for the cleanest representation of `Αναστολή θεραπείας`.
2. Implement orphan final outcomes: placement or date-specific suspension.
3. Integrate suspension into daily statistics.
4. Add focused tests for orphan resolution, auditability, and non-destructive recurring behavior.
5. Native-check orphan workflow.
6. Test selected/all restore with multiple changes.
7. Replace snapshot restore with authoritative overlay removal + recompute.
8. Resume outpatient recurring create/edit native gate.
9. Clean Python 3.14 ZipFile teardown warnings later.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Outpatients cannot receive Ρομποτικό.
- Nonblank manual constraints are never silently overwritten.
- Standard duration 45 minutes; Ανακλινόμενο 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative for inpatient recurring; `OUTPATIENT_SCHEDULE` authoritative for outpatient recurring.
- one-day cancellations, absences, replacements, suspensions and redistributions are overlays, not recurring edits,
- operational load/capacity is timeslot-based,
- existing displaced patients precede new patients,
- unresolved orphans must be placed or explicitly suspended,
- `MASTER` and `THERAPIST_DAILY` must agree on effective daily state.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
