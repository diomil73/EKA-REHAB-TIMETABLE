# Current Handoff

Last updated: 2026-10-18

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Trusted daily source for restore smoke copies: `Excel/previews/RERESOLVED_V90.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- focused gate after MASTER batch-write patch: `35 passed in 3.67s`
- diagnostic restore-all before MASTER batch write: Backend `45.70s`
- diagnostic restore-all after MASTER batch write: Backend `31.33s`, rollback `0.15s`, preview open `0.88s`, sheet sync `0.19s`, final save `1.23s`
- direct backend profile after fact-session consolidation: total `22.924s`, pre-Excel/verification `12.978s`, Excel output `9.946s`
- `restore_fact_batch 8.863s` has been eliminated from the pipeline
- OOXML MASTER projection path is regression-green: `59 passed in 6.85s`
- latest direct backend profile after OOXML projection migration: total `16.509s`, pre-Excel/verification `7.976s`, Excel output `8.533s`
- latest `build_master_projection`: `0.089s` (down from `6.142s`)
- latest Excel output detail: dispatch `1.347s`, open `1.921s`, MASTER `2.613s`, daily patches `0.452s`, statistics `0.316s`, replacement fact rows `0.160s`, ledger rows `0.024s`, refresh `0.039s`, save `1.390s`, close `0.066s`, quit `0.004s`
- parity rebuild: row/selected/all smoke copies authoritative and VBA-preserving
- per-row / selected / restore-all in-session correctness: native green
- replacement lifecycle history V87-V92: native green

## CLOSED: authoritative daily truth and replacement lifecycle

Concrete daily truth remains reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Replacement facts preserve immutable history. First placement keeps the legacy id, re-placement after restore gets `:v2`, later cycles get `:v3`, etc. Restore never rewrites a newer lifecycle fact.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a projection. The primary Excel restore path keeps the workbook open, launches Python hidden and blocking, creates rollback protection, imports the verified operational state in-session, and saves without close/reopen or visible console.

Native correctness is green for per-row, selected multi-row, and restore-all controls. User-facing `Αιτία` remains Greek-only for known replacement causes.

## DEFERRED: restore performance / UX latency

Restore latency has improved materially and is no longer the active engineering priority. Backend profiling is now `16.509s` on the trusted V90 scenario, with the MASTER metadata parser reduced from `6.142s` to `0.089s`. End-user restore time around the low-to-mid 20-second range is accepted for the current development stage.

Do not spend more time on restore micro-optimization now. Remaining latency work, including `pre_excel_unattributed`, Excel process startup/open/save overhead, and any further batching, is explicitly deferred until the broader feature set is substantially complete.

Performance work already completed and retained:

- selected/all lifecycle batching,
- statistics + ledger finalization consolidation,
- redundant in-session save removal,
- single restore output Excel session,
- MASTER projection moved into that output session,
- bounded meaningful-range sync,
- range-batched MASTER projection write-back,
- detailed backend profiler,
- reuse of already-loaded data,
- replacement fact status persistence folded into the existing output Excel session,
- OOXML package-probe metadata path for MASTER projection,
- OPC relationship target normalization for package-absolute and relative paths.

## Exact resume point

1. Do not resume restore performance tuning unless a new correctness regression appears or the broader project reaches final polish.
2. Resume the next unfinished functional milestone: outpatient recurring create/edit native verification and the wider scheduling/UI roadmap.
3. Preserve the current restore profiler and timings as the baseline for end-stage optimization.
4. Keep correctness gates locked: source/VBA preservation, immutable history, validate-all-before-mutate, exact statistics, and MASTER/THERAPIST_DAILY agreement.
5. Windows tests always use `python -m pytest`.

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

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
