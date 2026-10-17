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
- initial daily restore bridge preview-builder + ledger contract gate: **`16 passed in 2.94s`**
- fact-to-ledger synchronization + ledger contract rerun: **`18 passed in 4.40s`**
- rebuilt row/selected/all smoke copies after parity fix: **all three green with `Active ledger rows: 5`, fact sync=True, source unchanged=True, VBA preserved=True, bridge module refreshed=True**
- first native per-row button smoke: **backend preview verification reached success, but UI lifecycle failed**; visible Python console opened, then `ThisWorkbook.Close` failed with `Method 'Close' of object '_Workbook' failed`
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

The backend JSON response contract is source-green with `30 passed in 4.31s`, the complete ledger bridge source gate is green with **`42 passed in 4.27s`**, the fact-to-ledger synchronization gate is green with **`18 passed in 4.40s`**, and the rebuilt native smoke copies have exact authoritative ledger parity.

The first real Excel per-row button test exposed a UI lifecycle regression, not a daily-state/backend regression. The backend reached the verified-preview success message, but the bridge then attempted to transition through a detached commit-worker close/reopen cycle. Two problems were visible:

- `WScript.Shell.Exec(commandLine)` exposed a black Python console window.
- the bridge later executes `ThisWorkbook.Close SaveChanges:=True`; from the active Form Control macro this failed with `Method 'Close' of object '_Workbook' failed`.

The user also explicitly rejects a close/reopen UX for normal ledger actions. Selected/all native smoke tests are therefore paused.

### Historical precedent recovered

Git history confirms this lifecycle problem was already encountered in registration work:

- `86678fd29fa283a2485187a0edc0836875c86daa`: deferred authoritative commit worker introduced.
- `b06d3ea18fafe2b716e6066294280f63f27158a4`: worker-driven Excel close exposed.
- `0af80eecdb0bd0f0ed20f60e6fdf3d48596399af`: sentinel workbook workaround kept Excel alive while closing the authoritative workbook.
- `97691a1f44aefe621799ea2ae5507cc48fc7c51a`: direct COM close tried to simplify the lifecycle.
- `7b03cfd5a020f254734da6dde974bcbbeb01a1d8`: the proven sentinel/self-close method was restored only as a fallback; the code explicitly notes that registration's primary UI path no longer depends on it.
- `db67abb6487182c12c2634627b04f2972b70b912`: **Register patients directly inside the open workbook**. The current registration bridge and outpatient in-session writer save `ThisWorkbook` directly without closing/reopening it.

This is the pattern to restore for the daily ledger UX: Python remains authoritative for validation/recompute, but the final accepted result must be applied inside the already-open workbook and saved in-session. The detached close/reopen worker is not an acceptable primary ledger UI path.

### Projection parity blocker already fixed

The first three smoke copies had exposed a stale UI projection: only 1 ACTIVE ledger row versus 5 authoritative ACTIVE facts. The smoke-copy builder now treats `DAILY_REPLACEMENT_FACTS` as the ledger projection source. After the fix and the green 18-test source rerun, all three independent copies were rebuilt successfully:

- `LEDGER_ROW_SMOKE.xlsm`: ACTIVE ledger rows=5, synchronized=True, source unchanged=True, VBA preserved=True, bridge module refreshed=True
- `LEDGER_SELECTED_SMOKE.xlsm`: ACTIVE ledger rows=5, synchronized=True, source unchanged=True, VBA preserved=True, bridge module refreshed=True
- `LEDGER_ALL_SMOKE.xlsm`: ACTIVE ledger rows=5, synchronized=True, source unchanged=True, VBA preserved=True, bridge module refreshed=True

## Exact resume point

1. **Do not test `LEDGER_SELECTED_SMOKE.xlsm` or `LEDGER_ALL_SMOKE.xlsm` yet.**
2. Replace the ledger primary UI lifecycle with an in-session authoritative commit path: no `ThisWorkbook.Close`, no automatic reopen, no detached close worker.
3. Replace visible `WScript.Shell.Exec` backend launch with a hidden, blocking launch that still returns the JSON result/exit status.
4. Add source regressions proving the ledger path has no visible console launch and no workbook-close/reopen dependency.
5. Rebuild `LEDGER_ROW_SMOKE.xlsm` from trusted V90 and rerun only the per-row button smoke.
6. Require: workbook stays open, no console window, exact fact lifecycle/recompute, ledger status update, MASTER/THERAPIST_DAILY agreement, exact `DAILY_STATISTICS`, and VBA preservation.
7. Only after per-row native green, test selected two-fact restore and then restore-all on their independent copies.
8. Once all three are native-green, close the ledger migration milestone and resume outpatient recurring create/edit native gate.

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
