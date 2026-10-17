# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Trusted daily source for restore smoke copies: `Excel/previews/RERESOLVED_V90.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement fact focused gate: `28 passed, 3 warnings in 5.14s`
- versioned-id + single/batch restore focused gate: `36 passed in 4.74s`
- rerun gate after V90 persistence fix: `37 passed in 7.13s`
- restore JSON bridge/backend gate: `30 passed in 4.31s`
- ledger authoritative bridge source gate: `42 passed in 4.27s`
- initial restore bridge preview-builder gate: `16 passed in 2.94s`
- fact-to-ledger synchronization gate: `18 passed in 4.40s`
- focused bridge/performance gate after first batching patch: `30 passed in 4.91s`
- focused bridge/performance gate after statistics+ledger session consolidation: `30 passed in 4.39s`
- focused bridge/performance gate after redundant in-session saves were removed: `31 passed in 3.34s`
- focused gate after single output Excel session: `32 passed in 4.48s`
- focused gate after MASTER-session removal: `33 passed in 3.81s`
- focused gate after bounded meaningful-range sync: `34 passed in 2.31s`
- focused gate after timing instrumentation: `35 passed in 1.56s`
- diagnostic restore-all timing: Backend `45.70s`, rollback `0.14s`, preview open `0.88s`, sheet sync `0.20s`, final save `1.19s`
- parity rebuild: row/selected/all smoke copies reported authoritative ledger parity and preserved source/VBA
- in-session per-row restore: native green
- in-session selected two-row restore: native correctness green
- in-session restore-all: native correctness green
- replacement lifecycle history V87-V92: native green

## CLOSED: authoritative daily truth and replacement lifecycle

Concrete daily truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Replacement facts preserve immutable history. First placement keeps the legacy id, re-placement after restore gets `:v2`, later cycles get `:v3`, etc. Restore never rewrites a newer lifecycle fact.

Atomic selected/all backend restore is implemented and native green independently through V91/V92: validate the whole requested set before mutation, no partial restore, one fact mutation batch, one daily-state recompute, one statistics rewrite, source/VBA preservation.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a projection. The smoke-copy builder rebuilds the ledger from durable facts and preserves matching audit/snapshot extras by immutable ReplacementId.

The primary Excel restore path keeps the workbook open, launches Python hidden and blocking, creates a rollback copy, imports the verified operational state in-session, and saves the workbook. The obsolete visible-console and close/reopen lifecycle is not part of the normal path.

Native correctness is green for per-row, selected multi-row, and restore-all controls. Restore-all independently verified ACTIVE replacements 0, treatment suspensions 2, unresolved 5, finalized ΟΧΙ, exact stored statistics, preserved historical RESTORED facts, and no visible console.

User-facing `Αιτία` is Greek-only for known replacement causes. Technical/internal reason values such as `DAILY_INPUT therapist absence` and `temporary replacement` render as `Απουσία θεραπευτή`. Internal immutable IDs remain unchanged.

## ACTIVE: restore performance / UX

The diagnostic timing isolated the problem decisively: almost all latency is inside the authoritative Python/backend leg. Excel-side rollback/open/sync/save totals only about 2.4 seconds, while backend execution is about 45.7 seconds.

Performance chain already applied:

- selected/all ledger lifecycle batching
- statistics + ledger finalization consolidation
- redundant in-session save removal
- single restore output Excel session
- MASTER projection moved into that output session
- bounded meaningful-range sync instead of broad `UsedRange`
- explicit phase timing in the Excel success dialog

The next backend optimization is now implemented: MASTER projection write-back no longer performs row-by-row Copy/Paste/Clear and cell-by-cell formula assignment. It restores formatting with at most two range-level PasteSpecial operations, clears the projection body in one range operation, and writes the formula grid in one bulk `.Formula` assignment. Only responsible-doctor rich text and infectious styling remain per-row. Implementation `d1fdced7ee46d63277dffa372c14572858345dea`, regression `eb6942296fbf3668e066c35bf672bb18fcc4cd4c`.

## Exact resume point

1. Pull latest branch.
2. Run `python -m pytest Tests/test_authoritative_bulk_daily_restore.py Tests/test_daily_restore_in_session_vba.py Tests/test_daily_change_ledger_vba.py Tests/test_daily_restore_bridge_preview.py -q`.
3. If green, build a fresh restore-all smoke copy from trusted V90 and run one restore-all. Never reuse an already-restored smoke workbook.
4. Compare the new `Backend` timing with the previous `45.70s`; also record the remaining four phase timings.
5. If backend is still materially above target, instrument sub-phases inside `_RestoreOutputBackend` next, especially workbook open, MASTER batch write, daily patches, statistics/ledger refresh, and output save.
6. Once restore latency is acceptable, resume outpatient recurring create/edit native gate.

## Locked orphan rule

Every orphan for the concrete date resolves as exactly one of:

- valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`. Recurring/base remains untouched. A finalized day requires unresolved=0.

## Locked recomputation order

1. recurring/base schedule
2. cancellations/no-shows release capacity
3. therapist absences
4. identify displaced/orphan sessions
5. apply surviving ACTIVE replacements
6. resolve remaining orphans by placement or suspension
7. only then expose capacity to new patients

Operational capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Documentation discipline

Meaningful changes, regressions, native results, decisions, needs, planned additions/removals/fixes must be recorded in repository documentation during the same work cycle. Chat history is not the sole source of truth.

## Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cosmetic cleanup. Not a blocker.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
