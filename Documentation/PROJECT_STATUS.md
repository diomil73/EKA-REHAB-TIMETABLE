# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness is closed green for per-row, selected, and restore-all Excel paths. Performance/UX latency remains active. The dominant remaining cost is pre-Excel parsing/verification.**

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
- focused gate after MASTER batch-write patch: **35 passed in 3.67s**
- native diagnostic after MASTER batch write: **Backend 31.33s**, rollback 0.15s, preview open 0.88s, sheet sync 0.19s, final save 1.23s
- direct backend profile: **32.673s total**, **24.646s pre-Excel/verification**, **8.027s Excel output**
- Excel output detail: dispatch 1.398s, workbook open 1.356s, MASTER 1.787s, daily patches 0.415s, statistics 0.220s, ledger rows 0.019s, refresh 0.033s, save 1.626s, close 0.363s, quit 0.252s
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

MASTER projection batching produced a real gain: backend time fell from 45.70 seconds to 31.33 seconds. The next profiler run showed that only about 8.0 seconds now belong to the entire hidden Excel output session, while about 24.6 seconds occur before/around that session in parsing, reconstruction and verification. Excel-side apply after the backend remains about 2.4 seconds.

Implemented optimization chain:

1. Bulk selected/all ledger rows are marked in one hidden Excel session instead of one session per ReplacementId.
2. Daily statistics and ledger finalization share the same session.
3. Redundant success-path saves in the in-session bridge were removed while rollback and authoritative post-sync save remain.
4. Native daily write plan, statistics, restored ledger metadata, ledger refresh, and final output save share one `_RestoreOutputBackend` Excel session.
5. MASTER projection is derived in Python and rebuilt inside that already-open restore output workbook.
6. Broad `UsedRange` copies during verified-preview apply were replaced by bounded meaningful ranges resolved with Excel `Find`.
7. The Excel bridge reports backend, rollback, preview-open, sheet-sync, and final-save timings.
8. MASTER projection write-back is range-batched: one body clear, one 2-D formula-grid assignment, and at most two formatting PasteSpecial calls. Implementation `d1fdced7ee46d63277dffa372c14572858345dea`, regression `eb6942296fbf3668e066c35bf672bb18fcc4cd4c`.
9. The backend profiler now instruments both Excel-output phases and individual pre-Excel phases. Deeper instrumentation commit `6b28548e4a9a45c2ad8b59ffc5ea650db4fa4aca`, tests `836d4db2c5a568f3c1662244275bfcb9a3253bb6`.
10. The restore pipeline now reuses data it already computed instead of reopening/parsing the workbook for identical information: `read_unified_base_schedule(..., patients=patients)` avoids a duplicate PATIENTS load, and `build_master_daily_overlay_plan(..., master_projection_rows=master_projection)` avoids rebuilding MASTER projection a second time. Implementation commits `a14c75a4c2664f348be67d59c5ccdbbe1d2ef17e`, `6eac109d96462f9255ad041c8c1d8d5649eec4cb`, `874b71c381e0996b2970960ba67a49f17f0e9eb0`; regression commits `fc8cd34dfd074768f0723215036cd4cc16a61d13`, `d56eaadef3885645db4effa8b18f51c51b52c5c6`, `97f21934ae4b0e74af275ec89ee02f8c8af8f787`.

## Immediate roadmap

1. Pull latest branch and run the focused test set including profiler, MASTER overlay, and outpatient schedule source tests.
2. Rerun the backend profiler against trusted `RERESOLVED_V90.xlsm` using a fresh `BACKEND_PROFILE2.xlsm` output.
3. Use the newly exposed pre-Excel timings to optimize the largest remaining read/parser phase. Likely candidates are repeated workbook loads in patient/base/daily/replacement readers, MASTER projection, patient preview role scanning, or final hash/VBA verification.
4. Re-run fresh reconstruction after the next performance change.
5. Keep the acceptance target at roughly **<=15 seconds total**, ideally around 10 seconds.
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
