# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Excel ledger authoritative restore bridge, source gate pending after VBA wiring**.

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- versioned-id + single/batch restore focused gate: **36 passed in 4.74s**
- rerun gate after V90 persistence fix: **37 passed in 7.13s**
- restore JSON bridge/backend gate: **30 passed in 4.31s**
- V87 mixed orphan scenario: **native green**
- V88 full restart reconstruction: **native green**
- V89 single authoritative restore + reconstruction: **native green**
- V90 re-placement + reconstruction + versioned lifecycle: **native green**
- V91 selected authoritative restore + fresh reconstruction: **native green**
- V92 restore-all + fresh reconstruction: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

The operational day is reconstructable from durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

V88 proved this across a fresh process with 5 replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ, and exact `DAILY_STATISTICS` equivalence.

## CLOSED: versioned replacement lifecycle

Replacement facts preserve immutable history across restore/re-placement cycles:

- original generation keeps the legacy id
- re-placement after restore becomes `:v2`
- later cycles become `:v3`, etc.

V90 proved this natively: the original planner:57 fact remained `RESTORED`, the later placement became a distinct `:v2` ACTIVE fact, and fresh reconstruction remained exact.

## CLOSED: atomic selected/all restore

Implemented rules:

- validate the entire requested restore set before mutation
- any missing/stale/already-restored member aborts the complete batch
- no partial restore
- selected restore by repeated exact ids
- restore-all over all ACTIVE facts
- one authoritative mutation batch
- one daily-state recompute
- one statistics rewrite
- preserve source workbook and VBA

Key files:

- `Python/rehab_excel/daily_replacement_fact_batch.py`
- `Python/tools/restore_daily_replacement_facts.py`
- `Tests/test_daily_replacement_fact_batch.py`
- `Tests/test_authoritative_bulk_daily_restore.py`

## V91 selected restore result

`Excel/previews/SELECTED_RESTORE_V91.xlsm` restored exactly two ACTIVE replacement facts from V90.

Native result:

- restored replacement facts=2
- active replacements after restore=3
- unresolved sessions after recompute=2
- finalized after recompute=ΟΧΙ
- source unchanged=True
- VBA preserved=True

Fresh reconstruction independently confirmed 3 active replacements, 2 treatment suspensions, unresolved=2, finalized=ΟΧΙ, and exact `DAILY_STATISTICS` equivalence.

## V92 restore-all result

`Excel/previews/RESTORE_ALL_V92.xlsm` was created independently from trusted V90 using `--all-active`.

Native result:

- restored replacement facts=5
- active replacements after restore=0
- unresolved sessions after recompute=5
- finalized after recompute=ΟΧΙ
- source unchanged=True
- VBA preserved=True

Fresh reconstruction independently confirmed persisted replacements=0, treatment suspensions=2, unresolved=5, finalized=ΟΧΙ, and exact stored-statistics equivalence.

## ACTIVE: Excel ledger authoritative bridge

The restore CLI now has a JSON response contract for Excel/VBA and its backend-focused gate is green at **30 passed in 4.31s**.

The ledger VBA has now been rewired to authoritative semantics:

- per-row sends exactly one immutable ReplacementId
- selected sends one atomic request containing all selected exact ids
- restore-all uses one `--all-active` request
- each action calls `restore_daily_replacement_facts.py` once
- the JSON response returns verified preview path + source SHA
- the existing authoritative commit worker performs the final safe replace only after Excel releases the source workbook
- the workbook reopens on `THERAPIST_DAILY`
- direct VBA snapshot cell-copy restore statements are removed; snapshots remain audit context only

New source-test contract also explicitly forbids the legacy `Cells(...).Copy Destination:=...` restore path.

Relevant commits:

- JSON bridge response contract: `2c5a67388479f840d064172108bb61a38d1640fa`
- JSON bridge tests: `749c9f50625ae44ce247829a236b06a4e4244afd`
- VBA bridge contract tests: `5c2269547752eedc923c741a1ed99a8786e06894`
- VBA bridge implementation: `9f1b9449d4d4822c6e095ec2438405bc9e71f5c3`

## Immediate roadmap

1. Pull latest branch.
2. Run the ledger bridge source gate after the VBA change.
3. If green, build a fresh native preview containing the new ledger module.
4. Smoke-test per-row restore from Excel.
5. Smoke-test selected two-fact restore from Excel.
6. Smoke-test restore-all from Excel on a separate trusted copy.
7. For each path verify fact lifecycle, fresh daily-state recompute, ledger status, MASTER/THERAPIST_DAILY agreement, `DAILY_STATISTICS`, source integrity, and VBA preservation.
8. Once native-green, close the ledger migration milestone and resume outpatient recurring create/edit native gate.
9. Later clean Python 3.14 ZipFile warnings.

## Locked daily rules

1. recurring/base schedule
2. cancellations/no-shows release capacity
3. therapist absences
4. identify displaced/orphan sessions
5. apply surviving ACTIVE replacements
6. unresolved orphans resolve through placement or treatment suspension
7. only then expose capacity to new patients

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Documentation discipline

Every meaningful evolution must be recorded in repository documentation during the same work cycle, including implemented changes, regressions, native validation results, newly discovered necessities, planned additions, removals, replacements, and deferred fixes. Chat history is not the sole source of truth.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Safety constraints

- never write directly to baseline `.xlsm` during development
- preview workflows use copies
- preserve and verify VBA
- keep layout write zones narrow
- existing replacement facts must never be re-appended during rerun
- partial preview files created before a later persistence failure must be removed
- bulk restore validates the whole set before any mutation
- Excel ledger restore actions must mutate authoritative fact lifecycle, never directly copy operational cells back
- Windows tests use `python -m pytest`

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
