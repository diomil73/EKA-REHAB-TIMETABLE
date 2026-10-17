# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Excel ledger authoritative restore bridge. Per-row native path is green with hidden Python + in-session Excel apply. Selected and restore-all native gates remain.**

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- versioned-id + single/batch restore focused gate: **36 passed in 4.74s**
- rerun gate after V90 persistence fix: **37 passed in 7.13s**
- restore JSON bridge/backend gate: **30 passed in 4.31s**
- ledger authoritative bridge source gate: **42 passed in 4.27s**
- bridge preview-builder gate: **16 passed in 2.94s**
- fact-to-ledger synchronization gate: **18 passed in 4.40s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row in-session Excel restore: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

Operational truth is reconstructable from durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

V88-V92 prove restart-safe reconstruction, immutable replacement lifecycle, atomic selected/all backend semantics, and exact stored-statistics equivalence.

## ACTIVE: Excel ledger authoritative bridge

### Projection contract

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a UI/audit projection, never an independent truth source. The builder synchronizes the ledger from durable facts and preserves matching audit extras by immutable ReplacementId.

### UI lifecycle correction

The first real Excel button test exposed two unacceptable lifecycle behaviors in the original bridge: visible `WScript.Shell.Exec` console and `ThisWorkbook.Close` failure. Git history showed that registration had already solved the same class of problem by moving the primary workflow into the already-open workbook.

The ledger now follows that pattern:

- authoritative Python restore/recompute still produces a verified preview
- backend launch is hidden and blocking
- current workbook remains open
- rollback copy is created first
- verified operational sheets are synchronized into the open workbook
- workbook is saved in-session
- no primary close/reopen worker
- user success message contains no implementation wording

### Per-row native result

A fresh `LEDGER_ROW_SMOKE.xlsm` was built from trusted `RERESOLVED_V90.xlsm` with 5 ACTIVE ledger rows and exact authoritative parity. One ACTIVE row was restored from Excel without closing the workbook.

Fact inventory after the action:

- total facts = 6
- RESTORED = 2
- ACTIVE = 4
- restored row = `planner:33...`
- unrelated `planner:57...:v2` remained ACTIVE

Fresh reconstruction independently confirmed:

- persisted replacements = 4
- treatment suspensions = 2
- unresolved = 1
- finalized = ΟΧΙ
- stored `DAILY_STATISTICS` exactly matches reconstructed daily state

The **per-row native gate is closed green**.

Relevant recent commits:

- in-session bridge implementation: `446f6f599a30f4a561b4be8a53967d1e02703e69`
- builder wiring: `b7fdb206c2e9e3b3ed5758beb52e64d7056f5a54`
- source regression updates: `8776b690f53d3b4972f87ad311da9efc863173dd`, `61a2f703936ea73dd1834095cf15e9ecf671a7f4`, `762b67dd086ad496f0c9cf3ac08927f955657eea`
- success-message cleanup: `b1554ba1c5588ad936e704e5c91ab04cfdca306f`, `70cad9d1d3a44ea065a9a75da91202208518ab07`

## Immediate roadmap

1. Pull latest branch.
2. Rebuild independent `LEDGER_SELECTED_SMOKE.xlsm` and `LEDGER_ALL_SMOKE.xlsm` from trusted V90 so they contain the current in-session bridge.
3. Selected native test: restore exactly two ACTIVE rows atomically. Expect active replacements 5 -> 3, unresolved=2, finalized=ΟΧΙ, suspensions=2, exact statistics.
4. Verify selected result with fresh reconstruction and replacement inventory.
5. Restore-all native test on its independent copy. Expect active replacements 5 -> 0, unresolved=5, finalized=ΟΧΙ, suspensions=2, exact statistics.
6. Require no visible console, no workbook close/reopen, correct ledger history, MASTER/THERAPIST_DAILY agreement, source/VBA safety.
7. Once selected + all are native green, close the ledger migration milestone and resume outpatient recurring create/edit native gate.
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

## Safety constraints

- never write directly to baseline `.xlsm` during development
- preview workflows use copies
- preserve and verify VBA
- keep layout write zones narrow
- existing replacement facts must never be re-appended during rerun
- partial previews created before later failure must be removed
- bulk restore validates the whole set before any mutation
- Excel ledger restore mutates authoritative fact lifecycle, never old snapshot cells directly
- `DAILY_CHANGE_LEDGER` remains a projection of authoritative facts
- normal ledger UI must not close/reopen the workbook or expose a console window
- Windows tests use `python -m pytest`

## Documentation discipline

Every meaningful evolution must be recorded in repository documentation during the same work cycle. Chat history is not the sole source of truth.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
