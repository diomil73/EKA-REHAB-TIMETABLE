# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness is closed green for per-row, selected, and restore-all Excel paths. The active ledger item is performance/UX latency.**

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- versioned-id + single/batch restore focused gate: **36 passed in 4.74s**
- rerun gate after V90 persistence fix: **37 passed in 7.13s**
- restore JSON bridge/backend gate: **30 passed in 4.31s**
- ledger authoritative bridge source gate: **42 passed in 4.27s**
- bridge preview-builder gate: **16 passed in 2.94s**
- fact-to-ledger synchronization gate: **18 passed in 4.40s**
- focused bridge/performance gate after first batching patch: **30 passed in 4.91s**
- focused bridge/performance gate after statistics+ledger consolidation: **30 passed in 4.39s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row in-session Excel restore: **native green**
- selected two-row in-session Excel restore: **native correctness green**
- restore-all in-session Excel restore: **native correctness green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

Operational truth is reconstructable from durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

V88-V92 prove restart-safe reconstruction, immutable replacement lifecycle, atomic selected/all backend semantics, and exact stored-statistics equivalence.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a UI/audit projection, never an independent truth source.

The normal Excel restore workflow now:

- runs the authoritative Python backend hidden and blocking
- keeps the current workbook open
- creates a rollback copy before apply
- imports the verified operational state back into the open workbook
- saves in-session
- avoids the obsolete visible-console and close/reopen path
- shows Greek-only business wording in the visible `Αιτία` column

Native correctness is green for per-row, selected, and restore-all controls.

Restore-all independently confirmed the expected end state:

- active replacements = 0
- treatment suspensions = 2
- unresolved = 5
- finalized = ΟΧΙ
- stored `DAILY_STATISTICS` exactly matches reconstructed daily state
- replacement history remains preserved as RESTORED facts

The Greek reason rendering patch maps known internal values such as `DAILY_INPUT therapist absence` and `temporary replacement` to `Απουσία θεραπευτή` in the normal visible ledger.

## ACTIVE: restore performance / UX latency

The remaining ledger issue is speed. Selected and restore-all actions take noticeably longer than acceptable.

Three optimization steps are implemented:

1. Bulk selected/all restore no longer reopens Excel once per restored ReplacementId merely to mark ledger rows. All restored ledger rows are validated and marked in a single hidden Excel session, followed by one ledger refresh/save. Implementation `fb445bbc08e6d2cbfd63820a13f8dc504192e954`, tests `de7dbe61a8bae19e469cbe93c309a6fcdf18d4c7`.
2. Daily statistics writing and ledger finalization now share that same hidden Excel session, removing another workbook open/save/close cycle. Implementation `d92dda3b82b9d12ea32e63592cfe80110f5af22c`, tests `7b38e29996a17557190984f5ab9763382e1ab33e`.
3. The in-session apply path now removes two redundant success-path writes: it no longer re-saves the workbook immediately before `SaveCopyAs`, because the source was already saved before backend execution, and it no longer performs an extra ledger refresh/save after the verified preview has already been synchronized and saved. Rollback `SaveCopyAs`, authoritative preview synchronization, and the post-sync save remain. Implementation `4b58085fc416c456c603aefd1ae2942022715632`, tests `55b7e584080d6f9f657ce3b274dae9a90a85aa19`.

Focused source gate after the second optimization: **30 passed in 4.39s**. The third optimization still needs the focused Windows source gate and one fresh native timing smoke.

Further profiling is still required. Remaining likely hotspots are:

- MASTER projection Excel open/save
- native write-plan Excel open/save
- verified preview opening and broad `UsedRange` copying in the already-open workbook

Optimization must preserve authoritative verification, rollback safety, immutable fact history, and exact reconstructed statistics.

## Immediate roadmap

1. Pull latest branch and run the focused bridge/performance test set.
2. Rebuild a fresh selected/all smoke copy from trusted V90. Do not reuse a workbook that has already been restored.
3. Repeat the native restore action without an extra manual save and compare elapsed time.
4. Run fresh reconstruction after the action to prove no correctness regression.
5. If latency remains high, consolidate the MASTER projection and native write-plan Excel phases next, then narrow the in-session synchronization contract.
6. Once latency is acceptable, resume outpatient recurring create/edit native gate.
7. Later clean Python 3.14 ZipFile warnings.

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
- normal ledger UI must not expose internal English reason labels in `Αιτία`
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
