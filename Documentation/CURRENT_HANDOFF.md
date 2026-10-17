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
- native restore-all timing after those patches: about `50-60s`, still unacceptable
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

The primary Excel restore path now keeps the workbook open, launches Python hidden and blocking, creates a rollback copy, imports the verified operational state in-session, and saves the workbook. The obsolete visible-console and close/reopen lifecycle is not part of the normal path.

Native correctness is green for all three Excel controls:

- per-row restore
- selected multi-row restore
- restore-all

Restore-all was verified from an independent V90 copy with the expected final state:

- ACTIVE replacements: 0
- treatment suspensions: 2
- unresolved sessions: 5
- finalized: ΟΧΙ
- stored `DAILY_STATISTICS` matches fresh reconstruction
- historical replacement facts remain auditable as RESTORED
- workbook stayed open during the action, with no visible console

User-facing `Αιτία` is Greek-only for known replacement causes. Technical/internal reason values such as `DAILY_INPUT therapist absence` and `temporary replacement` render as `Απουσία θεραπευτή`. Internal immutable IDs remain unchanged.

## ACTIVE: restore performance / UX

Correctness is closed. The remaining ledger item is latency. The latest fresh restore-all smoke still takes about 50-60 seconds, so performance is far from the acceptance target.

Five performance patches are now on the branch:

- Bulk selected/all restore marks all restored ledger rows in one hidden Excel session instead of reopening Excel once per ReplacementId. Implementation `fb445bbc08e6d2cbfd63820a13f8dc504192e954`, tests `de7dbe61a8bae19e469cbe93c309a6fcdf18d4c7`.
- Daily statistics write and ledger finalization are consolidated into the same hidden Excel session. Implementation `d92dda3b82b9d12ea32e63592cfe80110f5af22c`, tests `7b38e29996a17557190984f5ab9763382e1ab33e`.
- In-session apply removes redundant success-path saves while preserving rollback and post-sync save. Implementation `4b58085fc416c456c603aefd1ae2942022715632`, tests `55b7e584080d6f9f657ce3b274dae9a90a85aa19`.
- Native daily write plan, statistics, restored-ledger lifecycle marking, ledger refresh, and output save share one `_RestoreOutputBackend` Excel session. Implementation `790f55babd18a75635d844cc4f6d1f24c2c5780e`, tests `0c166fff3e3dc26fd8b53062653e9764234d2b98`. Focused gate: `32 passed in 4.48s`.
- New patch `0c2e566614679fa010ffec09fb52825eea17055b` removes the separate MASTER projection Excel process. MASTER projection is now derived in Python and rebuilt inside the already-open `_RestoreOutputBackend` workbook before the daily patches, statistics and ledger lifecycle are applied. Regression commit `ac08df5a617fd159ba40f196299b86242a2546a3`.

Remaining likely hotspots after the MASTER-session removal are the verified-preview open inside the user's Excel session and the broad `UsedRange` synchronization across the restore sheet set.

## Exact resume point

1. Pull latest branch.
2. Run `python -m pytest Tests/test_authoritative_bulk_daily_restore.py Tests/test_daily_restore_in_session_vba.py Tests/test_daily_change_ledger_vba.py Tests/test_daily_restore_bridge_preview.py -q`.
3. If green, rebuild a fresh `LEDGER_ALL_SMOKE_FAST2.xlsm` from trusted V90. Never reuse an already-restored smoke workbook.
4. Run restore-all without a manual extra save and time from click to success message.
5. Run fresh reconstruction to prove the MASTER-session consolidation did not alter authoritative state.
6. If latency remains above ~15s, next target is the in-session verified-preview open plus broad `UsedRange` synchronization, ideally replacing broad sheet copies with a narrow verified patch contract.
7. Once restore latency is acceptable, resume outpatient recurring create/edit native gate.

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
