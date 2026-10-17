# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness is closed green for per-row, selected, and restore-all Excel paths. Performance/UX latency remains active and is still unacceptable.**

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
- focused bridge/performance gate after redundant in-session saves were removed: **31 passed in 3.34s**
- focused gate after single output Excel session: **32 passed in 4.48s**
- focused gate after MASTER-session removal: **33 passed in 3.81s**
- focused gate after bounded meaningful-range sync: **34 passed in 2.31s**
- focused gate after timing instrumentation: **35 passed in 1.56s**
- diagnostic restore-all timing: **Backend 45.70s**, rollback 0.14s, preview open 0.88s, sheet sync 0.20s, final save 1.19s
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

The normal Excel restore workflow runs the authoritative Python backend hidden and blocking, keeps the current workbook open, creates rollback protection, imports verified operational state, saves in-session, avoids visible console/close-reopen behavior, and renders Greek-only business wording in the visible `Αιτία` column.

Native correctness is green for per-row, selected, and restore-all controls. Restore-all independently confirmed active replacements 0, treatment suspensions 2, unresolved 5, finalized ΟΧΙ, exact stored statistics, and preserved RESTORED history.

## ACTIVE: restore performance / UX latency

The diagnostic timing now isolates the dominant cost: backend generation is about 45.7 seconds, while all Excel-side apply phases after the backend total only about 2.4 seconds. Optimization therefore moves inside the backend rather than further changing rollback/preview synchronization.

Implemented optimization chain:

1. Bulk selected/all ledger rows are marked in one hidden Excel session instead of one session per ReplacementId.
2. Daily statistics and ledger finalization share the same session.
3. Redundant success-path saves in the in-session bridge were removed while rollback and authoritative post-sync save remain.
4. Native daily write plan, statistics, restored ledger metadata, ledger refresh, and final output save share one `_RestoreOutputBackend` Excel session.
5. MASTER projection is derived in Python and rebuilt inside that already-open restore output workbook.
6. Broad `UsedRange` copies during verified-preview apply were replaced by bounded meaningful ranges resolved with Excel `Find`.
7. The Excel bridge reports backend, rollback, preview-open, sheet-sync, and final-save timings.
8. New backend optimization: MASTER projection write-back is now range-batched. The projection body is cleared once, formulas are assigned as one 2-D `.Formula` matrix, and formatting uses at most two range-level PasteSpecial calls. Per-row COM traffic is retained only for responsible-doctor rich text and infectious highlighting. Implementation `d1fdced7ee46d63277dffa372c14572858345dea`, regression `eb6942296fbf3668e066c35bf672bb18fcc4cd4c`.

## Immediate roadmap

1. Pull latest branch and run the focused bridge/performance tests.
2. Build a fresh restore-all smoke copy from trusted V90 and run restore-all once.
3. Compare the new Backend timing against the previous **45.70s** reference.
4. If backend remains materially above the target, add sub-phase timings inside `_RestoreOutputBackend` for workbook open, MASTER write, daily patches, statistics/ledger refresh, and save, then optimize the measured dominant sub-phase.
5. Re-run fresh reconstruction after performance changes.
6. Once restore latency is acceptable, resume outpatient recurring create/edit native gate.
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
