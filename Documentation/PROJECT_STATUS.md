# Project Status

Last updated: 2026-10-17

## Current stage

The branch remains `feature/outpatient-scheduling`.

The active operational milestone is now **orphan resolution + date-specific treatment suspension + correct daily statistics**, built on top of the already working concrete-date effective state shared by MASTER, THERAPIST_DAILY and the ledger.

Latest checkpoints:

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Current native checkpoint

`Excel/previews/DAILY_CHANGE_LEDGER_V85.xlsm` is the latest native checkpoint.

Confirmed native green for the ΜΕΝΤΗΣ stress case:

- active replacement appears consistently in THERAPIST_DAILY and MASTER,
- source trace remains visible,
- individual restore returns both views to original,
- time axis displays `hh:mm` with no `####`,
- moved non-robotic ΜΕΝΤΗΣ no longer inherits robotic-orange styling.

## Deferred cosmetic issue

The obsolete `ΑΝΑΝΕΩΣΗ` Form Control remains visible.

It has been identified exactly as `Button 1` with caption `ΑΝΑΝΕΩΣΗ` and macro `RefreshAll`. Multiple targeted cleanup attempts did not remove it reliably. Decision: **defer this issue and do not spend more engineering time on it now**. It is not a functional blocker; the user may remove it manually in Excel later.

## Active orphan workflow

Every patient in `Ορφανοί` for the concrete date must resolve as either:

- placed in a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Required behavior:

- no silent disappearance,
- suspension is a daily overlay/status, never a recurring delete,
- suspension is auditable,
- suspension feeds daily statistics as treatment not performed with explicit reason,
- existing displaced patients precede new-patient allocation,
- recomputation remains idempotent and independent of input order.

## Daily operational rules locked

1. recurring/base schedule for the concrete date,
2. cancellations/no-shows release same-day capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Operational capacity remains timeslot-based. Recurring/base schedules are never destructively edited by one-day facts.

## Ledger status

Native green:

- active move rendering,
- MASTER sync,
- source trace,
- per-row restore,
- time-axis display,
- stale-orange cleanup.

Still open:

- selected/all restore with multiple changes,
- replacing preview snapshot restore with authoritative overlay removal + full daily recomputation.

## Outpatient architecture

The branch already contains inpatient/outpatient types, separate authoritative `OUTPATIENT_SCHEDULE`, unified base reads, outpatient presentation, cancellation/no-show overlays and recurring schedule infrastructure.

Outpatient recurring create/edit native verification remains open after the current orphan/daily-state milestone.

## Immediate roadmap

1. Inspect the existing daily-state, cancellation and statistics architecture.
2. Implement date-specific `Αναστολή θεραπείας` and orphan final outcomes.
3. Integrate the suspension reason into daily statistics.
4. Add focused tests for auditability, idempotence and non-destructive recurring behavior.
5. Native-check the orphan workflow.
6. Test selected/all restore with multiple changes.
7. Replace snapshot restore with authoritative overlay removal + full recompute.
8. Resume outpatient recurring create/edit native gate.
9. Continue branch/PR hardening.
10. Clean Python 3.14 ZipFile teardown warnings later.

## Safety constraints

- Never write directly to the baseline `.xlsm` during development or smoke testing.
- Preview workflows operate on copied `.xlsm` files only.
- Preserve/verify VBA when rebuilding previews.
- Keep current scheduling behavior backward-compatible unless an explicit business rule changes it.
- Windows test command is always `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
