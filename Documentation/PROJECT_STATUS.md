# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness is closed green for per-row, selected, and restore-all Excel paths. Performance/UX latency remains active. The two largest measured backend costs have now been redesigned and await Windows timing verification.**

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- focused gate after MASTER batch-write patch: **35 passed in 3.67s**
- native diagnostic after MASTER batch write: **Backend 31.33s**, rollback 0.15s, preview open 0.88s, sheet sync 0.19s, final save 1.23s
- direct backend profile after read-reuse work: **28.794s total**, **20.911s pre-Excel/verification**, **7.883s Excel output**
- largest measured phases: **restore_fact_batch 8.863s**, **build_master_projection 5.833s**, **pre_excel_unattributed 2.667s**
- Excel output detail: dispatch 1.528s, workbook open 1.268s, MASTER 1.612s, daily patches 0.540s, statistics 0.271s, ledger rows 0.022s, refresh 0.037s, save 1.666s, close 0.202s, quit 0.048s
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row / selected / restore-all in-session correctness: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

Operational truth is reconstructable from durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

V88-V92 prove restart-safe reconstruction, immutable replacement lifecycle, atomic selected/all semantics, and exact stored-statistics equivalence.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a UI/audit projection, never an independent truth source.

The normal Excel restore workflow runs the authoritative Python backend hidden and blocking, keeps the current workbook open, creates rollback protection, imports verified operational state, saves in-session, avoids visible console/close-reopen behavior, and renders Greek-only business wording in the visible `Αιτία` column.

Native correctness is green for per-row, selected, and restore-all controls.

## ACTIVE: restore performance / UX latency

Profiling showed that Excel-side preview synchronization is no longer the bottleneck. The dominant backend costs were a second native Excel process used only to flip replacement fact statuses and repeated openpyxl loads inside MASTER projection.

Implemented optimization chain now includes:

1. selected/all ledger lifecycle batching,
2. statistics + ledger finalization consolidation,
3. redundant in-session save removal,
4. single restore output Excel session,
5. MASTER projection moved into that output session,
6. bounded meaningful-range sync instead of broad `UsedRange`,
7. explicit phase timing in the Excel success dialog,
8. range-batched MASTER projection write-back,
9. detailed backend profiling,
10. reuse of already-loaded PATIENTS and already-built MASTER projection,
11. **new:** replacement fact lifecycle persistence is moved into the existing `_RestoreOutputBackend` Excel session. The staging workbook is no longer opened/saved in a second Excel process. Requested ids are validate-all-first in memory, effective replacements are derived by excluding those ids, and the verified output copy receives `RESTORED` statuses before its single save. Implementation `c7895e28e5843df30a889b55326189764c3dec0a`, regressions `140126a6a3f09ccaa9fe4b29f8a09547e8408119` and `9a6e7243c7a02d3191b5feedbf59de4d45addfb2`.
12. **new:** MASTER projection metadata now reads PATIENTS row ids, SETTINGS room order, and MASTER header location in one openpyxl workbook session rather than three. Implementation `63e6920c700cd73644d93a9fe8ffd79c365105cc`, regression `44c8e6556849fc4a091179a43dae7962e52cf5b5`.
13. Profiler updated for the new lifecycle and `replacement_fact_rows` output phase: `4d928b884e81264620b9c666016a77f23aa0e96e`.

Safety semantics are unchanged: the complete requested restore set is validated before any authoritative status mutation. Python still does not save the staging `.xlsm`; native Excel mutates only the output preview copy.

## Immediate roadmap

1. Pull latest branch and run the focused tests including authoritative restore, profiler, MASTER projection/overlay, outpatient schedule source, and in-session bridge tests.
2. Rerun `profile_daily_restore_backend.py` against trusted `RERESOLVED_V90.xlsm` using a fresh `BACKEND_PROFILE3.xlsm`.
3. Confirm that the old `restore_fact_batch` cost has disappeared and MASTER projection time has dropped materially.
4. Optimize the next largest measured phase immediately. Likely next target is the non-Excel portion of `apply_write_plan_total` / copy-verification overhead, followed by remaining workbook parser loads.
5. Re-run fresh reconstruction after performance changes.
6. Keep the acceptance target at roughly **<=15 seconds total**, ideally around 10 seconds.
7. Once restore latency is acceptable, resume outpatient recurring create/edit native gate.
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
