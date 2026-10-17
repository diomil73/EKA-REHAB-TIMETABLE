# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest trusted native restore preview: `Excel/previews/AUTHORITATIVE_RESTORE_V89.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement load + mixed outpatient/inpatient: `31 passed in 2.04s`
- MASTER daily overlay + ledger restore source gate: `24 passed, 5 warnings in 4.39s`
- orphan suspension + statistics domain gate: `21 passed in 3.11s`
- consolidated orphan operational gate: `43 passed, 3 warnings in 5.15s`
- replacement fact focused gate: `28 passed, 3 warnings in 5.14s`
- versioned-id + single/batch restore focused gate: **36 passed in 4.74s**
- rerun gate after V90 persistence fix: **37 passed in 7.13s**
- V87 mixed orphan resolution: native green
- V88 full restart daily-fact reconstruction: native green
- V89 single authoritative restore + fresh reconstruction: **native green**

Known Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe daily truth

`ORPHAN_RESOLUTION_V88.xlsm` fresh-process reconstruction confirmed:

- persisted replacements reread: 5
- treatment suspensions reread: 2
- unresolved: 0
- finalized: ΝΑΙ
- stored `DAILY_STATISTICS` exactly matches reconstructed state

Concrete daily truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

## CLOSED: first authoritative restore proof

`AUTHORITATIVE_RESTORE_V89.xlsm` restored exactly:

`daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19`

Native inventory confirmed:

- restored fact remains as audit history with `RESTORED`,
- other 4 replacement facts remain `ACTIVE`.

Fresh reconstruction confirmed:

- persisted replacements reread: 4
- treatment suspensions reread: 2
- unresolved: 1
- finalized: ΟΧΙ
- stored `DAILY_STATISTICS` exactly matches reconstructed state

This proves restore now changes authoritative fact lifecycle and recomputes the day instead of copying old cells. Reopened orphan after restore is expected because the therapist absence still exists.

## ACTIVE: versioned IDs + selected/all authoritative restore

### Versioned replacement IDs implemented

Operational placement reads full replacement history, including RESTORED rows.

ID rule:

- first generation keeps legacy id: `daily-input:<session_id>:<date>`
- next placement after restore becomes `...:v2`
- later cycles become `...:v3`, etc.

This prevents a stale restore from targeting a newer replacement for the same session/date.

Commits:

- regression: `faeb6c3b455479731cad615cf9075a89524ab4f8`
- implementation: `9fc9799e5e8942bafb6af9b37333e8cf02a2b198`

### Atomic selected/all restore backend implemented

New module:

- `Python/rehab_excel/daily_replacement_fact_batch.py`

Rules:

- validates the complete requested set before any mutation,
- missing/stale/already-restored member stops the whole batch,
- no partial restore,
- only after full validation are all requested facts marked RESTORED and saved once.

New tool:

- `Python/tools/restore_daily_replacement_facts.py`

Supports:

- repeated `--replacement-id` for selected restore,
- `--all-active` for restore all,
- one fact mutation batch,
- one daily-state recompute,
- one statistics rewrite,
- MASTER recurring reset + surviving overlays,
- source/VBA preservation.

Commits:

- batch backend: `c2fc50a8e7f16166415ab4ff7371c3cccfbbb7ac`
- batch validation tests: `2d3afa3ef081fb2c368fb59c53224e34a79f8c28`
- selected/all tool: `d109af5eed7b09354e3e13bdf8e065d58ba817db`
- selected/all pipeline tests: `b8690432a5927de671e5e08edc41e3a57bc0d5af`

### 2026-10-17 focused gate regression and resolution

The first Windows focused source gate after the bulk-restore pull produced `35 passed, 1 failed` because `Tests/test_daily_fact_reconstruction_verifier.py` still called `_replacement_id(session)` after the API became history-aware.

The test was corrected to `_replacement_id(session, [])` for the no-history/base-id case.

Fix commit:

- `c43b677667c59d38b65b928d4126c36cf670f28c`

Rerun result:

- **36 passed in 4.74s**

### V90 re-placement regression discovered and fixed

The first native V89 -> V90 re-placement attempt stopped with:

`SAFETY STOP: Duplicate ReplacementId 'daily-input:planner:61:ΦΘ@2026-10-19:2026-10-19'`

Root cause:

- `apply_daily_input_replacements.py` correctly reread the 4 surviving ACTIVE replacement facts from V89,
- but after resolving the reopened orphan it passed the complete active replacement list back to `append_daily_replacement_facts`,
- therefore the writer attempted to append the 4 already-persisted facts again before it could add only the new `:v2` fact.

The duplicate-id guard behaved correctly; the caller was wrong.

Fix implemented:

- existing ACTIVE replacements remain in memory for workload/state reconstruction,
- only replacements created during the current run are appended as new facts,
- post-copy persistence failures remove the newly created preview so a half-written preview is not left behind,
- output reporting distinguishes replacements loaded from disk, replacements created this run, and total active replacements after resolution.

Commits:

- implementation/safety cleanup: `ae0f58d363ec0cbf357876cf6b4633e3f14d70e7`
- regression coverage: `88b75319273005b28d373d2a683701f82fe41fba`

Focused rerun after the fix:

- **37 passed in 7.13s**

Important: any `RERESOLVED_V90.xlsm` left by the failed attempt is untrusted/partial and must not be inspected as authoritative evidence. The corrected rerun uses `--overwrite` and must recreate V90 from V89.

## Exact resume point

1. Re-run V89 -> V90 with `apply_daily_input_replacements.py`; choose a candidate, not suspension.
2. Inspect V90 replacement facts; require the original `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19` as `RESTORED` and new `...:v2` as `ACTIVE`.
3. Fresh reconstruction verifier must match stored statistics with expected replacements=5, suspended=2, unresolved=0.
4. Native test selected restore of two ACTIVE facts with `restore_daily_replacement_facts.py`; expected 2 restored, 3 active, unresolved=2, finalized=ΟΧΙ.
5. Native test `--all-active` from V90 on a separate preview; expected 5 restored, 0 active, unresolved=5, finalized=ΟΧΙ.
6. Connect Excel ledger per-row/selected/all buttons to authoritative bridge path; only then retire old cell-copy VBA restore implementation.
7. Resume outpatient recurring create/edit native gate.

## Locked orphan rule

Every orphan for the concrete date must resolve as exactly one of:

- valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`.
Recurring/base remains untouched. A finalized day requires unresolved=0.

## Locked recomputation order

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. resolve remaining orphans by placement or suspension,
7. only then expose capacity to new patients.

Operational capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Documentation discipline

New developments, regressions, decisions, necessities, deferred fixes, planned removals and newly agreed features must be captured in repository documentation during the same work cycle. Chat history is not an acceptable sole record.

## Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cosmetic cleanup. Not a blocker.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
