# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **authoritative restore hardening + selected/all restore**.

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics domain gate: **21 passed in 3.11s**
- consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- V87 mixed orphan scenario: **native green**
- V88 full restart reconstruction: **native green**
- V89 single authoritative restore + reconstruction: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

V88 proves the operational day survives restart through durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Fresh reconstruction produced 5 replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ and exact DAILY_STATISTICS equivalence.

## CLOSED: first authoritative restore

V89 restored one exact V88 fact through lifecycle change and full recompute.

Confirmed after fresh reread:

- 1 historical fact = RESTORED,
- 4 replacement facts = ACTIVE,
- treatment suspensions=2,
- unresolved=1,
- finalized=ΟΧΙ,
- stored statistics exactly match reconstructed state.

The reopened orphan is expected because the original therapist absence remains active.

## ACTIVE: hardening and bulk restore

Versioned replacement IDs are implemented:

- first generation keeps the legacy id,
- re-placement after restore becomes `:v2`,
- later cycles become `:v3`, etc.

Atomic batch restore is also implemented:

- validate entire selected set before mutation,
- missing/stale/already-restored member aborts the complete batch,
- no partial restore,
- selected restore by repeated exact ids,
- restore-all over all ACTIVE facts,
- one daily-state recompute and one statistics rewrite.

Files:

- `Python/rehab_excel/daily_replacement_fact_batch.py`
- `Python/tools/restore_daily_replacement_facts.py`
- `Tests/test_daily_replacement_fact_batch.py`
- `Tests/test_authoritative_bulk_daily_restore.py`

Status: **implementation complete; focused Windows gate pending**.

## Immediate roadmap

1. Pull latest branch and run focused tests for version IDs, single restore and batch restore.
2. Re-resolve the one reopened V89 orphan with `apply_daily_input_replacements.py`; require new fact id ending `:v2`.
3. Verify resulting workbook reconstructs with unresolved=0, finalized=ΝΑΙ and exact statistics equivalence.
4. Test selected authoritative restore of two ACTIVE facts into a new preview.
5. Test `--all-active` restore on a separate preview.
6. Connect Excel ledger per-row/selected/all controls to authoritative backend; retire old cell-copy VBA only after native bridge gate.
7. Resume outpatient recurring create/edit native gate.
8. Later clean Python 3.14 ZipFile warnings.

## Locked daily rules

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. unresolved orphans resolve through placement or treatment suspension,
7. only then expose capacity to new patients.

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Safety constraints

- never write directly to baseline `.xlsm` during development,
- preview workflows use copies,
- preserve and verify VBA,
- keep layout write zones narrow,
- bulk restore validates the whole set before any mutation,
- Windows tests use `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
