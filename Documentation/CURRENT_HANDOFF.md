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
- current focused in-session bridge gate: `25 passed in 2.13s`
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

Correctness is closed. The remaining ledger item is latency: selected and restore-all actions take noticeably longer than acceptable for normal use.

Do not weaken the authoritative preview/recompute/rollback guarantees merely to mask the delay. Profile and optimize the expensive stages. Prime suspects:

- full Python preview/reconstruction pipeline
- opening the verified preview workbook in Excel
- copying full `UsedRange` blocks for multiple operational sheets
- repeated saves before/after apply

The goal is materially faster native restore while preserving the same verified final state and rollback safety.

## Exact resume point

1. Pull the latest branch if needed.
2. Profile the in-session restore path and identify where elapsed time is spent.
3. Reduce unnecessary workbook I/O and over-broad sheet synchronization without weakening authoritative verification or rollback.
4. Re-run the focused bridge tests and one native restore smoke after each performance change.
5. Once restore latency is acceptable, resume the outpatient recurring create/edit native gate.

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
