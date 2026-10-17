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
- versioned-id + single/batch restore focused gate: **36 passed in 4.74s**
- rerun gate after V90 persistence fix: **37 passed in 7.13s**
- V87 mixed orphan scenario: **native green**
- V88 full restart reconstruction: **native green**
- V89 single authoritative restore + reconstruction: **native green**
- V90 re-placement preview creation: **native green**
- V90 fresh reconstruction verification: **native green**
- V90 replacement-fact lifecycle inventory: **native green**

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

The focused source gate is green after the V90 rerun-persistence fix:

- **37 passed in 7.13s**

The earlier V89 -> V90 attempt exposed duplicate re-append of already persisted ACTIVE facts. That was fixed so only replacements created in the current run are appended, while existing ACTIVE facts remain in memory for workload/state reconstruction. Partial preview cleanup was also added.

Commits:

- `ae0f58d363ec0cbf357876cf6b4633e3f14d70e7` — rerun persistence + partial-preview cleanup
- `88b75319273005b28d373d2a683701f82fe41fba` — regression coverage

## V90 native re-placement, reconstruction and fact lifecycle result

`Excel/previews/RERESOLVED_V90.xlsm` was recreated from V89 successfully using a candidate placement.

Confirmed from the native run:

- accepted replacements this run: 1
- active replacements after resolution: 5
- persisted replacements loaded: 4
- persisted replacement facts written: 1
- treatment suspensions created this run: 0
- unresolved sessions: 0
- scheduled=80
- performed=78
- replaced=5
- suspended=2
- source unchanged=True
- VBA preserved=True

Fresh reconstruction then confirmed independently from the saved workbook:

- persisted replacements reread: 5
- treatment suspensions reread: 2
- reconstructed unresolved: 0
- finalized=ΝΑΙ
- stored `DAILY_STATISTICS` matches reconstructed daily state exactly

Explicit inventory then confirmed 6 historical replacement-fact rows:

- original `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19` = `RESTORED`,
- four surviving unrelated facts = `ACTIVE`,
- new `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19:v2` = `ACTIVE`.

This validates the intended immutable lifecycle: restored history remains auditable and a later re-placement gets a new versioned id instead of reusing or resurrecting the old fact.

Current status: **all V90 preconditions are green. Proceed to selected restore V91.**

## Immediate roadmap

1. Restore exactly two ACTIVE facts from V90 into V91: the new `planner:57...:v2` fact and `planner:61...`; expect 2 restored, 3 active, unresolved=2, finalized=ΟΧΙ.
2. Test `--all-active` restore from V90 into V92; expect 5 restored, 0 active, unresolved=5, finalized=ΟΧΙ.
3. Connect Excel ledger per-row/selected/all controls to authoritative backend; retire old cell-copy VBA only after native bridge gate.
4. Resume outpatient recurring create/edit native gate.
5. Later clean Python 3.14 ZipFile warnings.

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
- existing replacement facts must never be re-appended during a rerun,
- partial preview files created before a later persistence failure must be removed,
- bulk restore validates the whole set before any mutation,
- Windows tests use `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
