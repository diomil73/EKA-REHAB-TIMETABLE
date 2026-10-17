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
- latest fresh restore-all native timing: **about 50 seconds**, still far above target
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

The normal Excel restore workflow now runs the authoritative Python backend hidden and blocking, keeps the current workbook open, creates rollback protection, imports verified operational state, saves in-session, avoids visible console/close-reopen behavior, and renders Greek-only business wording in the visible `Αιτία` column.

Native correctness is green for per-row, selected, and restore-all controls. Restore-all independently confirmed active replacements 0, treatment suspensions 2, unresolved 5, finalized ΟΧΙ, exact stored statistics, and preserved RESTORED history.

## ACTIVE: restore performance / UX latency

The latest native restore-all timing is about 50 seconds. This is not acceptable for routine operation, so optimization continues rather than treating the current path as finished.

Implemented optimization chain:

1. Bulk selected/all ledger rows are marked in one hidden Excel session instead of one session per ReplacementId. Implementation `fb445bbc08e6d2cbfd63820a13f8dc504192e954`, tests `de7dbe61a8bae19e469cbe93c309a6fcdf18d4c7`.
2. Daily statistics and ledger finalization share the same session. Implementation `d92dda3b82b9d12ea32e63592cfe80110f5af22c`, tests `7b38e29996a17557190984f5ab9763382e1ab33e`.
3. Redundant success-path saves in the in-session bridge were removed while rollback and authoritative post-sync save remain. Implementation `4b58085fc416c456c603aefd1ae2942022715632`, tests `55b7e584080d6f9f657ce3b274dae9a90a85aa19`.
4. Native daily write plan, statistics, restored ledger metadata, ledger refresh, and final output save now share one `_RestoreOutputBackend` Excel session. Implementation `790f55babd18a75635d844cc4f6d1f24c2c5780e`, tests `0c166fff3e3dc26fd8b53062653e9764234d2b98`; focused gate **32 passed in 4.48s**.
5. MASTER projection is derived in Python and rebuilt inside the already-open restore output workbook, removing a separate Excel open/save/close phase. Implementation `0c2e566614679fa010ffec09fb52825eea17055b`, regression `ac08df5a617fd159ba40f196299b86242a2546a3`; focused gate **33 passed in 3.81s**.
6. New optimization `973094f4808dc3bd42b83499468b245c24d6dbde` replaces broad `UsedRange` copies during in-session verified-preview apply. Each restore sheet now finds the last meaningful formula/value row and column and copies only that bounded rectangle, retaining Excel-native formatting and rich-text transfer while avoiding inflated formatting-only tails. Regression coverage: `c456da245fc3df3fe35b1917551c01a5a15cc742`.

The sixth optimization still needs the focused Windows test gate and a fresh native timing smoke.

## Immediate roadmap

1. Pull latest branch and run the focused bridge/performance tests.
2. Rebuild a fresh restore-all smoke copy from trusted V90 and re-time the native action. Do not reuse an already-restored workbook.
3. Re-run fresh reconstruction after the timing smoke.
4. If native latency remains above roughly 15 seconds, add explicit phase timing for backend generation, rollback copy, preview opening, sheet synchronization, and final save before the next structural optimization.
5. Once restore latency is acceptable, resume outpatient recurring create/edit native gate.
6. Later clean Python 3.14 ZipFile warnings.

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
