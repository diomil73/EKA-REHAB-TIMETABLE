# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **authoritative daily operational facts + restart-safe recomputation**, following successful native orphan resolution.

Latest checkpoints:

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics domain gate: **21 passed in 3.11s**
- consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**
- V87 mixed orphan scenario: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Native V87 result

`Excel/previews/ORPHAN_RESOLUTION_V87.xlsm`

Confirmed terminal/native result:

- 5 accepted replacement placements,
- 2 treatment suspensions,
- 2 persisted suspension facts,
- unresolved=0,
- scheduled=80,
- performed=78,
- replaced=5,
- suspended=2,
- `DAILY_STATISTICS` written,
- input unchanged=True,
- VBA preserved=True.

The V87 layout-safety blocker caused by stale outpatient cleanup outside the therapist grids was fixed without widening write zones.

## Persistence status

Suspensions are now durable daily facts in output `DAILY_INPUT` and the reader recognizes them as `TREATMENT_SUSPENDED`.

Replacement placements are **not yet durable authoritative facts**. Therefore full effective-state reconstruction after process/workbook restart is not complete yet.

New verification tool:

- `Python/tools/verify_orphan_resolution_persistence.py`
- `Tests/test_verify_orphan_resolution_persistence.py`

This tool verifies saved suspension facts against stored DAILY_STATISTICS and explicitly does not claim replacement reread support.

## Immediate roadmap

1. Verify V87 reread: 2 suspensions, statistics suspended=2, unresolved=0, finalized=ΝΑΙ.
2. Persist accepted replacement assignments as authoritative date-specific facts.
3. Rebuild full daily effective state from durable facts and prove restart/recompute equivalence.
4. Replace snapshot ledger restore with authoritative fact removal + recompute.
5. Test selected/all restore with multiple simultaneous changes.
6. Resume outpatient recurring create/edit native gate.
7. Continue branch hardening and later clean Python 3.14 ZipFile warnings.

## Locked daily rules

1. recurring/base schedule for concrete date,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced existing patients,
5. redistribute into net free capacity,
6. remaining orphans resolve by placement or treatment suspension,
7. new patients only after existing displaced patients are resolved.

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Recompute must be idempotent and order-independent.

## Existing native green areas

- MASTER toolbar/header UX,
- patient-centric scheduling,
- MASTER + THERAPIST_DAILY daily move sync,
- source trace,
- individual ledger restore,
- readable time axis,
- non-robotic color semantics,
- mixed orphan placement/suspension flow.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Safety constraints

- Never write directly to baseline `.xlsm` during development.
- Preview workflows use copies only.
- Preserve/verify VBA.
- Keep layout write zones narrow.
- Windows tests use `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
