# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest trusted native preview: `Excel/previews/SELECTED_RESTORE_V91.xlsm`
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
- V90 re-placement preview creation: **native green**
- V90 fresh daily-fact reconstruction: **native green**
- V90 replacement-fact lifecycle inventory: **native green**
- V91 selected authoritative restore: **native green**

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

This proves restore changes authoritative fact lifecycle and recomputes the day instead of copying old cells. Reopened orphan after restore is expected because the therapist absence still exists.

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

### V90 native re-placement, reconstruction and lifecycle inventory validated

`RERESOLVED_V90.xlsm` was recreated from trusted V89 with candidate placement, not suspension.

Native output confirmed:

- accepted replacements this run: 1
- active replacements after resolution: 5
- persisted replacements loaded: 4
- persisted replacement facts written: 1
- treatment suspensions created this run: 0
- unresolved sessions: 0
- statistics: scheduled=80, performed=78, replaced=5, suspended=2, unresolved=0
- source unchanged: True
- VBA preserved: True

Fresh reconstruction verifier confirmed 5 active replacements, 2 treatment suspensions, unresolved=0, finalized=ΝΑΙ, and exact `DAILY_STATISTICS` equivalence.

Explicit replacement-fact inventory confirmed 6 historical rows total:

- original `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19` remains `RESTORED`,
- four unrelated surviving replacement facts remain `ACTIVE`,
- new `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19:v2` is `ACTIVE`.

This closes the V90 versioned-id lifecycle proof: restore history is preserved while the new authoritative occurrence receives a distinct immutable id.

### V91 selected authoritative restore validated

`SELECTED_RESTORE_V91.xlsm` was created from V90 by restoring exactly two ACTIVE facts:

- `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19:v2`
- `daily-input:planner:61:ΦΘ@2026-10-19:2026-10-19`

Native output matched the expected atomic recompute result exactly:

- restored replacement facts: 2
- active replacements after restore: 3
- unresolved sessions after recompute: 2
- finalized after recompute: ΟΧΙ
- source unchanged: True
- VBA preserved: True

This is the first native proof that multi-select authoritative restore mutates the requested fact set together and recomputes the day from surviving truth rather than copying cells.

## Exact resume point

1. Fresh-verify `SELECTED_RESTORE_V91.xlsm`; expect persisted replacements=3, suspended=2, unresolved=2, finalized=ΟΧΙ, and exact stored-statistics equivalence.
2. Native test `--all-active` from V90 on a separate `RESTORE_ALL_V92.xlsm`; expected 5 restored, 0 active, unresolved=5, finalized=ΟΧΙ.
3. Fresh-verify V92 and inspect fact inventory if needed.
4. Connect Excel ledger per-row/selected/all buttons to authoritative bridge path; only then retire old cell-copy VBA restore implementation.
5. Resume outpatient recurring create/edit native gate.

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
