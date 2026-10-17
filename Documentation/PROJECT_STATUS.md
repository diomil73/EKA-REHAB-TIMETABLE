# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness is closed green for per-row, selected, and restore-all Excel paths. Performance/UX latency remains active. Backend is now down to about 23 seconds, the former 8.9-second fact-restore Excel process is gone, and the MASTER metadata path has moved to lightweight OOXML reads.**

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- focused gate after MASTER batch-write patch: **35 passed in 3.67s**
- native diagnostic after MASTER batch write: **Backend 31.33s**, rollback 0.15s, preview open 0.88s, sheet sync 0.19s, final save 1.23s
- latest direct backend profile: **22.924s total**, **12.978s pre-Excel/verification**, **9.946s Excel output**
- latest largest pre-Excel phase: **build_master_projection 6.142s**
- previous **restore_fact_batch 8.863s** phase is gone after lifecycle consolidation
- first OOXML projection gate exposed a package-absolute relationship-path bug: **5 failed, 51 passed in 5.06s**
- package relationship normalization fix + regression now green: **59 passed in 6.85s**
- latest Excel output: dispatch 2.133s, open 1.849s, MASTER 1.794s, daily patches 0.498s, statistics 0.277s, replacement facts 0.155s, ledger 0.174s, refresh 0.129s, save 2.189s, close 0.098s, quit 0.004s
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

The optimization chain now includes:

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
11. replacement fact lifecycle persistence moved into the existing output Excel session, eliminating the separate `8.863s` restore-fact Excel process,
12. MASTER metadata consolidated first to one workbook load, then further optimized because the full `.xlsm` load itself remained expensive,
13. `build_master_projection` now reads PATIENTS row ids, SETTINGS room order, and MASTER header directly through `WorkbookPackageProbe` from the OOXML ZIP/XML package. No full openpyxl load is used for these projection metadata reads. Implementation `82abf70349d52aaef627f7a83f5f2387d0b6f74a`, regression `5cf0fdfee7ad0c7d458b9b0e85c5e364c5a6271d`.
14. OPC relationship targets are normalized correctly for package-absolute (`/xl/...`), relative and parent-relative forms before ZIP access. Fix `0057aa38b9fc90b66a4bd5bef543a82edf5262b7`, regression `65dc8827dfa0d3f047d1823b68682f6232b2b395`. Focused gate is now `59 passed in 6.85s`.

Safety semantics are unchanged: the package probe is read-only, source bytes are not changed, validate-all-before-mutate remains intact, and native Excel still mutates only the output preview copy.

## Immediate roadmap

1. Run `profile_daily_restore_backend.py` against trusted `RERESOLVED_V90.xlsm` using fresh `BACKEND_PROFILE4.xlsm`.
2. Confirm that `build_master_projection` falls sharply from **6.142s**.
3. Optimize the next largest measured phase immediately. Likely targets are `pre_excel_unattributed ~2.58s` and then Excel process startup/open/save overhead.
4. Re-run fresh reconstruction after performance changes.
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
