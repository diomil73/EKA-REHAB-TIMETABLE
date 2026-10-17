# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest trusted native preview: `Excel/previews/RESTORE_ALL_V92.xlsm`
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
- V91 selected authoritative restore + fresh reconstruction: native green
- V92 restore-all + fresh reconstruction: **native green**

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

## CLOSED: selected/all authoritative restore

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

Fresh reconstruction independently confirmed 3 active replacements, 2 treatment suspensions, unresolved=2, finalized=ΟΧΙ, and exact `DAILY_STATISTICS` equivalence.

This closes the selected multi-fact restore proof.

### V92 restore-all fully validated

`RESTORE_ALL_V92.xlsm` was created independently from trusted V90 using `--all-active`.

Native authoritative restore output matched expectations exactly:

- restored replacement facts: 5
- active replacements after restore: 0
- unresolved sessions after recompute: 5
- finalized after recompute: ΟΧΙ
- source unchanged: True
- VBA preserved: True

Fresh reconstruction then independently confirmed from the saved V92 workbook:

- persisted replacements reread: 0
- treatment suspensions reread: 2
- reconstructed unresolved: 5
- reconstructed finalized: ΟΧΙ
- stored `DAILY_STATISTICS` exactly matches reconstructed daily state

This closes the native selected/all authoritative restore milestone. Both selected and restore-all paths now mutate authoritative replacement-fact lifecycle state, recompute from surviving truth, preserve immutable history, preserve VBA/source, and survive a fresh-process reconstruction check.

## ACTIVE: Excel ledger authoritative bridge

Next milestone is to connect the Excel replacement ledger controls to the authoritative backend rather than the old cell-copy restore path.

Required bridge behavior:

- per-row restore calls the authoritative restore backend for exactly that fact id
- selected restore sends the complete selected id set atomically
- restore-all targets the current ACTIVE set atomically
- after restore, refresh/recompute MASTER, daily overlays, ledger statuses, and `DAILY_STATISTICS` from authoritative truth
- preserve source workbook and VBA
- retire the old cell-copy VBA restore implementation only after the native bridge gate is green

### First bridge slice implemented, source gate pending

The existing VBA entry points were identified in `Python/rehab_excel/daily_change_ledger_vba.py`: `RestoreDailyChangeRow`, `RestoreSelectedDailyChanges`, and `RestoreAllDailyChanges`. The current private `RestoreLedgerRow` still performs the legacy direct cell-copy restore and must not be removed until the authoritative bridge is native-green.

The authoritative bulk restore CLI now exposes an optional JSON response contract for Excel/VBA via `--response`. Success responses include the preview path, source SHA-256, restore counts, unresolved/finalized state, and source/VBA preservation flags. Failure responses include `ok=false`, the error, and source SHA when available. The restore tool remains preview-only and still refuses source=output.

Commits:

- JSON bridge response contract: `2c5a67388479f840d064172108bb61a38d1640fa`
- focused bridge contract tests: `749c9f50625ae44ce247829a236b06a4e4244afd`

## Exact resume point

1. Pull the latest branch and run the focused bridge/backend source gate including `Tests/test_daily_restore_bridge_contract.py` and the existing authoritative restore tests.
2. If green, implement the VBA side: row/selected/all controls must collect exact immutable replacement ids, run `restore_daily_replacement_facts.py` once per requested batch, read the JSON response, then hand the preview + source SHA to the existing authoritative commit worker.
3. Add VBA source tests that require the bridge path and explicitly forbid the old `Cells(...).Copy Destination:=...` restore path once the new path is ready for native validation.
4. Build a new native preview and smoke-test per-row, selected, and restore-all from Excel itself.
5. Only after that native bridge gate is green, retire the old cell-copy restore implementation.
6. Resume outpatient recurring create/edit native gate afterwards.

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
