# Project Status

Last updated: 2026-10-17

## Current stage

The branch remains `feature/outpatient-scheduling`.

The active operational milestone is **orphan resolution + date-specific treatment suspension + correct daily statistics**, built on top of the already working concrete-date effective state shared by MASTER, THERAPIST_DAILY and the ledger.

Latest confirmed checkpoints:

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics focused source gate: **21 passed in 3.11s**

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

The obsolete `ΑΝΑΝΕΩΣΗ` Form Control remains visible. Decision: **defer/manual cleanup**. It is not a functional blocker and should not consume more engineering time now.

## Orphan workflow: domain foundation source-green

Every patient in `Ορφανοί` for the concrete date must resolve as either:

- placed in a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Architecture decision:

- `Αναστολή θεραπείας` uses `DailySessionStatus.CANCELLED` plus `SessionCancellationKind.TREATMENT_SUSPENDED`,
- no new parallel daily status is introduced,
- recurring/base schedule stays untouched,
- the provider slot is released,
- the explicit reason remains auditable.

Implemented source slice:

- `SessionCancellationKind.TREATMENT_SUSPENDED`,
- `rehab_core/orphan_resolution.py` with `create_treatment_suspension()`,
- exact reason `Αναστολή θεραπείας`,
- `rehab_core/daily_statistics.py`,
- statistics separate performed/replaced/absent/postponed/no-show/suspended/unresolved outcomes,
- unresolved `THERAPIST_ABSENT` sessions remain unresolved until placed or explicitly suspended,
- focused regression files added for suspension and statistics.

Status: **source green: 21 passed in 3.11s**. Excel operational actions and authoritative persistence are not wired yet.

## Daily operational rules locked

1. recurring/base schedule for the concrete date,
2. cancellations/no-shows release same-day capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Operational capacity remains timeslot-based. Recurring/base schedules are never destructively edited by one-day facts. Recompute must remain idempotent and order-independent.

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

1. Wire the actual orphan operational actions: place or suspend.
2. Persist suspension as an authoritative date-specific daily fact.
3. Surface daily statistics in workbook/reporting.
4. Native-check orphan placement and suspension.
5. Test selected/all restore with multiple changes.
6. Replace snapshot restore with authoritative overlay removal + full recompute.
7. Resume outpatient recurring create/edit native gate.
8. Continue branch/PR hardening.
9. Clean Python 3.14 ZipFile teardown warnings later.

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
