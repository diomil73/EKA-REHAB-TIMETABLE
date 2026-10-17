# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Excel ledger authoritative restore bridge; smoke copies parity-green, native Excel per-row button gate next**.

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- versioned-id + single/batch restore focused gate: **36 passed in 4.74s**
- rerun gate after V90 persistence fix: **37 passed in 7.13s**
- restore JSON bridge/backend gate: **30 passed in 4.31s**
- first ledger-VBA bridge gate attempt: **41 passed, 1 failed in 7.32s**
- ledger authoritative bridge source gate after assertion fix: **42 passed in 4.27s**
- initial daily restore bridge preview-builder + ledger contract gate: **16 passed in 2.94s**
- fact-to-ledger synchronization + ledger contract rerun: **18 passed in 4.40s**
- rebuilt row/selected/all smoke copies: **all three parity-green with 5 ACTIVE ledger rows, fact sync=True, source unchanged=True, VBA preserved=True, bridge module refreshed=True**
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

The restore CLI JSON response contract is green at **30 passed in 4.31s**, the complete post-wiring ledger bridge source gate is green at **42 passed in 4.27s**, the fact-to-ledger synchronization rerun is green at **18 passed in 4.40s**, and all three rebuilt smoke copies now match authoritative fact counts.

The ledger VBA is wired to authoritative semantics:

- per-row sends exactly one immutable ReplacementId
- selected sends one atomic request containing all selected exact ids
- restore-all uses one `--all-active` request
- each action calls `restore_daily_replacement_facts.py` once
- the JSON response returns verified preview path + source SHA
- the existing authoritative commit worker performs the final safe replace only after Excel releases the source workbook
- the workbook reopens on `THERAPIST_DAILY`
- direct VBA snapshot cell-copy restore statements are removed; snapshots remain audit context only

### Native projection regression and fix

The first three bridge copies exposed a stale UI projection: only 1 ACTIVE ledger row versus 5 authoritative ACTIVE facts. The builder was fixed so `DAILY_REPLACEMENT_FACTS` drives the ledger projection, including RESTORED history, while preserving any matching audit/snapshot extras by immutable ReplacementId.

The post-fix source gate passed **18/18**. The three smoke workbooks were then rebuilt from trusted `RERESOLVED_V90.xlsm` and all reported:

- `Active ledger rows: 5`
- `Ledger synchronized from authoritative facts: True`
- `Source unchanged: True`
- `VBA preserved: True`
- `Ledger bridge module refreshed: True`

This closes the parity blocker. The next gate is real Excel button execution, starting with the single-row restore path.

Relevant commits:

- authoritative fact-to-ledger synchronization: `618110bc2430ab96b3ebf365f1f68fe29c11b390`
- parity regression source tests: `1206c90f318cec4d0ea1bb0c63f49075b5acd37c`
- green post-fix source checkpoint docs: `1e1408562e3ff6656eb7d3e1e0e8f3247207aebb`

## Immediate roadmap

1. Open only `Excel/previews/LEDGER_ROW_SMOKE.xlsm` in native Excel.
2. On `DAILY_CHANGE_LEDGER`, run exactly one per-row `Επαναφορά` for an ACTIVE replacement.
3. Confirm the authoritative bridge completes, the workbook closes/reopens as designed, and the restored row becomes `RESTORED`.
4. Run fresh reconstruction/inventory checks on `LEDGER_ROW_SMOKE.xlsm` before touching the other two copies.
5. If green, run selected two-fact restore on `LEDGER_SELECTED_SMOKE.xlsm`.
6. Then run restore-all on `LEDGER_ALL_SMOKE.xlsm`.
7. For every path verify fact lifecycle, daily-state recomputation, ledger status, MASTER/THERAPIST_DAILY agreement, `DAILY_STATISTICS`, source integrity, and VBA preservation.
8. Once native-green, close the ledger migration milestone and resume outpatient recurring create/edit native gate.
9. Later clean Python 3.14 ZipFile warnings.

## Locked daily rules

1. recurring/base schedule
2. cancellations/no-shows release slots
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
- `DAILY_CHANGE_LEDGER` is a UI/audit projection of authoritative replacement facts, not an independent source of daily truth
- bridge smoke copies must be independent because a successful button action authoritatively replaces that smoke workbook
- Windows tests use `python -m pytest`

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
