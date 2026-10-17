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
- restore JSON bridge/backend gate: `30 passed in 4.31s`
- first ledger-VBA bridge gate attempt: `41 passed, 1 failed in 7.32s`
- ledger authoritative bridge source gate after assertion fix: **`42 passed in 4.27s`**
- daily restore bridge preview-builder + ledger contract gate: **`16 passed in 2.94s`**
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

### V92 restore-all fully validated

`RESTORE_ALL_V92.xlsm` was created independently from trusted V90 using `--all-active`.

Native authoritative restore output matched expectations exactly:

- restored replacement facts: 5
- active replacements after restore: 0
- unresolved sessions after recompute: 5
- finalized after recompute: ΟΧΙ
- source unchanged: True
- VBA preserved: True

Fresh reconstruction independently confirmed persisted replacements=0, treatment suspensions=2, unresolved=5, finalized=ΟΧΙ, and exact stored-statistics equivalence.

## ACTIVE: Excel ledger authoritative bridge

The backend JSON response contract is source-green with `30 passed in 4.31s`, the complete ledger bridge source gate is green with **`42 passed in 4.27s`**, and the dedicated smoke-copy builder gate is green with **`16 passed in 2.94s`**.

The VBA ledger bridge is implemented in `Python/rehab_excel/daily_change_ledger_vba.py`:

- per-row restore reads the exact immutable ReplacementId from ledger column 1 and sends one `--replacement-id`
- selected restore builds one command with repeated `--replacement-id` arguments, preserving atomic batch semantics
- restore-all sends one `--all-active`
- all three paths call `restore_daily_replacement_facts.py` exactly once for the requested operation
- the restore CLI produces a verified preview + source SHA via `--response`
- VBA then hands that preview to `authoritative_commit_worker_cli.py`
- the workbook closes, the verified preview is committed only after the source unlocks, then the authoritative workbook reopens on `THERAPIST_DAILY`
- the old direct snapshot cell-copy restore statements have been removed from VBA; stored snapshots remain audit context only

Commits:

- backend JSON bridge contract: `2c5a67388479f840d064172108bb61a38d1640fa`
- backend bridge tests: `749c9f50625ae44ce247829a236b06a4e4244afd`
- VBA authoritative bridge contract tests: `5c2269547752eedc923c741a1ed99a8786e06894`
- VBA authoritative bridge implementation: `9f1b9449d4d4822c6e095ec2438405bc9e71f5c3`
- selected-restore assertion compatibility fix: `3784c04a03f4ddebfa91cb73be19e8a7c1b926cc`
- native bridge preview builder: `f3d3dc4c2b2afa7943bb1fd243c59f469eebb092`
- preview-builder source tests: `6796a9a4453a5dc8b2c8a219e8bff7134ed1a798`

### Source gate closed

The first post-wiring gate returned `41 passed, 1 failed in 7.32s`. The sole failure was an assertion-text mismatch, not a behavior defect. After correcting the assertion to the valid explicit VBA `Call` syntax, the exact gate reran green at `42 passed in 4.27s`.

A dedicated Windows-native smoke-copy builder now exists at `Python/tools/build_daily_restore_bridge_preview.py`. It clones a trusted authoritative daily workbook, requires an existing `DAILY_CHANGE_LEDGER` with at least one ACTIVE row, refreshes only `modDailyChangeLedger` to the current bridge implementation, re-renders the ledger controls, verifies source immutability and VBA presence, and deletes a failed output copy. This avoids rebuilding V90 facts merely to test newer VBA. Its focused source gate is green at `16 passed in 2.94s`.

## Exact resume point

1. Build three independent smoke copies from trusted `RERESOLVED_V90.xlsm` using `build_daily_restore_bridge_preview.py`: one for per-row, one for selected, one for restore-all.
2. Smoke-test from Excel itself: per-row restore first, then selected two-fact restore, then restore-all on its separate copy.
3. For every native path require correct replacement-fact lifecycle, fresh recomputation, ledger statuses, MASTER/THERAPIST_DAILY agreement, `DAILY_STATISTICS`, source integrity, and VBA preservation.
4. Once native-green, mark the legacy cell-copy restore path retired and resume outpatient recurring create/edit native gate.

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
