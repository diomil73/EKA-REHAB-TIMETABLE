# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest trusted native preview: `Excel/previews/SELECTED_RESTORE_V91.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement fact focused gate: `28 passed, 3 warnings in 5.14s`
- versioned-id + single/batch restore focused gate: `36 passed in 4.74s`
- rerun gate after V90 persistence fix: `37 passed in 7.13s`
- V87 mixed orphan resolution: native green
- V88 restart-safe daily-fact reconstruction: native green
- V89 single authoritative restore + fresh reconstruction: native green
- V90 re-placement, fresh reconstruction, and lifecycle inventory: native green
- V91 selected authoritative restore: native green
- V91 fresh reconstruction: native green

Known Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe daily truth

Concrete daily truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

V88 proved restart-safe reconstruction with 5 replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ, and exact `DAILY_STATISTICS` equivalence.

## CLOSED: first authoritative restore proof

V89 restored exactly:

`daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19`

The restored fact remained as audit history with `RESTORED`; 4 unrelated replacement facts stayed `ACTIVE`. Fresh reconstruction gave 4 persisted replacements, 2 suspensions, unresolved=1, finalized=ΟΧΙ, and exact statistics equivalence.

## CLOSED: versioned replacement lifecycle

Replacement IDs are immutable per session/date history:

- first generation keeps `daily-input:<session_id>:<date>`
- next placement after restore becomes `...:v2`
- later cycles become `...:v3`, etc.

V90 proved the lifecycle natively:

- original planner:57 fact remains `RESTORED`
- four unrelated facts remain `ACTIVE`
- re-placement is new `planner:57...:v2` and `ACTIVE`
- V90 fresh reconstruction: 5 active replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ, exact statistics match

Relevant commits:

- versioned-id regression: `faeb6c3b455479731cad615cf9075a89524ab4f8`
- versioned-id implementation: `9fc9799e5e8942bafb6af9b37333e8cf02a2b198`
- regression API fix: `c43b677667c59d38b65b928d4126c36cf670f28c`
- rerun persistence / partial-preview cleanup: `ae0f58d363ec0cbf357876cf6b4633e3f14d70e7`
- rerun regression coverage: `88b75319273005b28d373d2a683701f82fe41fba`

## ACTIVE: selected/all authoritative restore

Atomic batch restore is implemented in:

- `Python/rehab_excel/daily_replacement_fact_batch.py`
- `Python/tools/restore_daily_replacement_facts.py`

Rules:

- validate the complete requested set before mutation
- missing/stale/already-restored member aborts the complete batch
- no partial restore
- selected restore accepts repeated exact ids
- `--all-active` restores every ACTIVE fact
- one fact mutation batch, one daily-state recompute, one statistics rewrite
- MASTER recurring reset + surviving overlays
- source/VBA preservation

Relevant commits:

- batch backend: `c2fc50a8e7f16166415ab4ff7371c3cccfbbb7ac`
- batch validation tests: `2d3afa3ef081fb2c368fb59c53224e34a79f8c28`
- selected/all tool: `d109af5eed7b09354e3e13bdf8e065d58ba817db`
- selected/all pipeline tests: `b8690432a5927de671e5e08edc41e3a57bc0d5af`

### V91 selected restore validated

`SELECTED_RESTORE_V91.xlsm` was created from V90 by restoring exactly:

- `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19:v2`
- `daily-input:planner:61:ΦΘ@2026-10-19:2026-10-19`

Native restore output:

- restored replacement facts: 2
- active replacements after restore: 3
- unresolved after recompute: 2
- finalized after recompute: ΟΧΙ
- source unchanged: True
- VBA preserved: True

Fresh reconstruction then independently confirmed:

- persisted replacements reread: 3
- treatment suspensions reread: 2
- reconstructed unresolved: 2
- reconstructed finalized: ΟΧΙ
- stored `DAILY_STATISTICS` exactly matches reconstructed daily state

This closes the selected multi-fact restore proof.

## Exact resume point

1. Run independent restore-all from trusted V90 into `RESTORE_ALL_V92.xlsm` using `--all-active`.
2. Expected V92 native result: 5 restored, 0 active, unresolved=5, finalized=ΟΧΙ, source unchanged=True, VBA preserved=True.
3. Fresh-verify V92; expect persisted replacements=0, suspended=2, unresolved=5, finalized=ΟΧΙ, exact stored-statistics equivalence.
4. Inspect V92 fact inventory if needed; all five formerly ACTIVE facts should now be `RESTORED`, including planner:57 `:v2`.
5. Connect Excel ledger per-row/selected/all buttons to the authoritative backend path; only then retire the old cell-copy VBA restore implementation.
6. Resume outpatient recurring create/edit native gate.

## Locked orphan rule

Every orphan for the concrete date resolves as exactly one of:

- valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`.
Recurring/base remains untouched. A finalized day requires unresolved=0.

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

New developments, regressions, decisions, necessities, deferred fixes, planned removals, and newly agreed features must be captured in repository documentation during the same work cycle. Chat history is not an acceptable sole record.

## Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cosmetic cleanup. Not a blocker.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
