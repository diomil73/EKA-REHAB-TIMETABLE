# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Excel ledger authoritative restore bridge; parity-green, but first native per-row button exposed an unacceptable close/reopen + visible-console lifecycle regression. In-session commit redesign is now active.**

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
- first native per-row button smoke: **backend preview reached verified-success state, but UI lifecycle failed** with visible Python console and `Method 'Close' of object '_Workbook' failed`
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

### Projection regression already fixed

The first three bridge copies exposed a stale UI projection: only 1 ACTIVE ledger row versus 5 authoritative ACTIVE facts. The builder was fixed so `DAILY_REPLACEMENT_FACTS` drives the ledger projection, including RESTORED history, while preserving any matching audit/snapshot extras by immutable ReplacementId.

The post-fix source gate passed **18/18**. The three smoke workbooks were then rebuilt from trusted `RERESOLVED_V90.xlsm` and all reported:

- `Active ledger rows: 5`
- `Ledger synchronized from authoritative facts: True`
- `Source unchanged: True`
- `VBA preserved: True`
- `Ledger bridge module refreshed: True`

### Native per-row lifecycle failure

The first real Excel per-row button test did not fail in replacement validation/recomputation. It reached the message that the restore preview had been verified and was ready for commit. The failure was the bridge's UI/process lifecycle:

- `WScript.Shell.Exec(commandLine)` opened a visible black Python console window.
- the bridge then attempted the detached commit-worker workflow and executes `ThisWorkbook.Close SaveChanges:=True`.
- the close attempt failed with `Method 'Close' of object '_Workbook' failed`.
- the user explicitly requires the normal restore action to keep the workbook open and not flash a console window.

Selected/all native tests are paused until this is corrected.

### Recovered historical solution

Git history shows the same class of workbook-lifecycle problem was solved earlier during registration work. The evolution was:

- deferred worker and worker-driven close were introduced (`86678fd...`, `b06d3ea...`)
- sentinel/self-close and direct COM-close variants were tried (`0af80eec...`, `97691a1f...`)
- `7b03cfd5...` restored the proven close fallback but explicitly documented that the **primary registration UI no longer depends on it**
- `db67abb6487182c12c2634627b04f2972b70b912` moved registration to **direct writes inside the already-open workbook**

The current `patient_registration_bridge_vba.py` and `outpatient_in_session_vba.py` follow that in-session pattern: suspend events/screen updates, mutate the open workbook through Excel/VBA, save `ThisWorkbook`, restore application state. No normal close/reopen cycle is required.

The ledger bridge must now follow the same UX principle while retaining Python as the authoritative validator/recompute engine. The detached close/reopen commit worker remains a fallback/recovery mechanism, not the primary ledger action.

Relevant bridge commits to date:

- JSON bridge response contract: `2c5a67388479f840d064172108bb61a38d1640fa`
- JSON bridge tests: `749c9f50625ae44ce247829a236b06a4e4244afd`
- VBA bridge contract tests: `5c2269547752eedc923c741a1ed99a8786e06894`
- VBA bridge implementation: `9f1b9449d4d4822c6e095ec2438405bc9e71f5c3`
- selected-restore assertion compatibility fix: `3784c04a03f4ddebfa91cb73be19e8a7c1b926cc`
- native bridge preview builder: `f3d3dc4c2b2afa7943bb1fd243c59f469eebb092`
- preview-builder source tests: `6796a9a4453a5dc8b2c8a219e8bff7134ed1a798`
- authoritative fact-to-ledger synchronization: `618110bc2430ab96b3ebf365f1f68fe29c11b390`
- parity regression source tests: `1206c90f318cec4d0ea1bb0c63f49075b5acd37c`

## Immediate roadmap

1. Stop selected/all native button testing until the per-row lifecycle is corrected.
2. Add regressions forbidding visible `Shell.Exec`, `ThisWorkbook.Close`, and primary reliance on `authoritative_commit_worker_cli.py` in the ledger path.
3. Implement hidden backend invocation and an **in-session authoritative apply/save** path for the verified restore result.
4. Rebuild only `LEDGER_ROW_SMOKE.xlsm` from trusted V90.
5. Repeat one per-row restore and require: no console, workbook remains open, exact fact lifecycle/recompute, ledger status, MASTER/THERAPIST_DAILY agreement, exact `DAILY_STATISTICS`, source/VBA safety.
6. Only after that is native-green, rebuild/test selected two-fact restore and restore-all independently.
7. Once all three are native-green, close the ledger migration milestone and resume outpatient recurring create/edit native gate.
8. Later clean Python 3.14 ZipFile warnings.

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
- Excel ledger restore actions must mutate authoritative fact lifecycle, never directly copy old snapshot cells back
- `DAILY_CHANGE_LEDGER` is a UI/audit projection of authoritative replacement facts, not an independent source of daily truth
- normal ledger UI must not require workbook close/reopen and must not expose a console window
- Windows tests use `python -m pytest`

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
