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

## Current validation state

The first focused Windows gate after the latest bulk-restore pull returned **35 passed, 1 failed**.

The failure was not in the backend restore logic. It was a stale regression call in `Tests/test_daily_fact_reconstruction_verifier.py`, which still invoked `_replacement_id(session)` although the implementation now requires `history_rows` to allocate stable versioned ids.

The regression was corrected to use `_replacement_id(session, [])` for the base/no-history case.

Fix commit:

- `c43b677667c59d38b65b928d4126c36cf670f28c`

Current status: **focused Windows source gate pending rerun; V90 must not be created until that gate is green**.

## Immediate roadmap

1. Pull latest branch and rerun focused tests for version IDs, single restore and batch restore.
2. Require a green source gate before any workbook mutation.
3. Re-resolve the one reopened V89 orphan with `apply_daily_input_replacements.py`; choose a candidate and require new fact id ending `:v2`.
4. Verify resulting V90 reconstructs with 5 replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ and exact statistics equivalence.
5. Test selected authoritative restore of two ACTIVE facts into V91; expected 2 restored, 3 active, unresolved=2, finalized=ΟΧΙ.
6. Test `--all-active` restore from V90 into V92; expected 5 restored, 0 active, unresolved=5, finalized=ΟΧΙ.
7. Connect Excel ledger per-row/selected/all controls to authoritative backend; retire old cell-copy VBA only after native bridge gate.
8. Resume outpatient recurring create/edit native gate.
9. Later clean Python 3.14 ZipFile warnings.

## Locked daily rules

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. unresolved orphans resolve through placement or treatment suspension,
7. only then expose capacity to new patients.

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Documentation discipline

Every meaningful evolution must be recorded in the repository documentation during the same work cycle, including implemented changes, regressions, native validation results, newly discovered necessities, planned additions, removals, replacements and deferred fixes. Chat history is not the sole source of truth.

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
