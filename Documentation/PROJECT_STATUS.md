# Project Status

Last updated: 2026-10-18

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness remains closed green. Restore performance has improved enough for now and is deferred; focus returns to unfinished functional work, beginning with outpatient recurring create/edit native verification and the broader scheduling/UI roadmap.**

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- focused gate after MASTER batch-write patch: **35 passed in 3.67s**
- package relationship normalization + OOXML projection gate: **59 passed in 6.85s**
- latest direct backend profile: **16.509s total**, **7.976s pre-Excel/verification**, **8.533s Excel output**
- `build_master_projection`: **0.089s**, down from **6.142s**
- previous **restore_fact_batch 8.863s** phase remains eliminated
- latest Excel output: dispatch 1.347s, workbook open 1.921s, MASTER 2.613s, daily patches 0.452s, statistics 0.316s, replacement facts 0.160s, ledger 0.024s, refresh 0.039s, save 1.390s, close 0.066s, quit 0.004s
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

## DEFERRED: restore performance / UX latency

The optimization chain has reduced the trusted restore backend from about `45.70s` initially to `16.509s` in the latest direct profile. The MASTER projection parser, once a `6.142s` hotspot, is now `0.089s` through the read-only OOXML package path.

The current end-user restore duration in the low-to-mid 20-second range is acceptable for this stage of development. Further performance work is explicitly postponed until the broader functional roadmap is much closer to completion.

Completed performance work retained in the codebase:

1. selected/all ledger lifecycle batching,
2. statistics + ledger finalization consolidation,
3. redundant in-session save removal,
4. single restore output Excel session,
5. MASTER projection rebuilt inside that session,
6. bounded meaningful-range sync instead of broad `UsedRange`,
7. phase timing instrumentation,
8. range-batched MASTER projection write-back,
9. detailed backend profiling,
10. reuse of already-loaded PATIENTS and already-built MASTER projection,
11. replacement fact lifecycle persistence folded into the existing output Excel session,
12. MASTER metadata migrated to direct OOXML package reads,
13. OPC relationship target normalization for package-absolute and relative paths.

Remaining latency sources such as `pre_excel_unattributed`, Excel process startup/open/save overhead, and further parser consolidation are deferred, not forgotten.

## Immediate roadmap

1. Resume outpatient recurring create/edit native verification.
2. Continue the broader scheduling and Excel UI/functionality roadmap instead of spending more cycles on restore micro-optimization.
3. Keep the current profiler and latest timings as the baseline for end-stage performance tuning.
4. Revisit restore performance near project completion or sooner only if a correctness/performance regression appears.
5. Later clean Python 3.14 ZipFile warnings.

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
