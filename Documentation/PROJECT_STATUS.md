# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **authoritative restore hardening + selected/all restore**.

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics domain gate: **21 passed in 3.11s**
- consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- versioned-id + single/batch restore focused gate: **36 passed in 4.74s**
- V87 mixed orphan scenario: **native green**
- V88 full restart reconstruction: **native green**
- V89 single authoritative restore + reconstruction: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

V88 proves the operational day survives restart through durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Fresh reconstruction produced 5 replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ and exact DAILY_STATISTICS equivalence.

## CLOSED: first authoritative restore

V89 restored one exact V88 fact through lifecycle change and full recompute.

Confirmed after fresh reread:

- 1 historical fact = RESTORED,
- 4 replacement facts = ACTIVE,
- treatment suspensions=2,
- unresolved=1,
- finalized=ΟΧΙ,
- stored statistics exactly match reconstructed state.

The reopened orphan is expected because the original therapist absence remains active.

## ACTIVE: hardening and bulk restore

Versioned replacement IDs are implemented:

- first generation keeps the legacy id,
- re-placement after restore becomes `:v2`,
- later cycles become `:v3`, etc.

Atomic batch restore is also implemented:

- validate entire selected set before mutation,
- missing/stale/already-restored member aborts the complete batch,
- no partial restore,
- selected restore by repeated exact ids,
- restore-all over all ACTIVE facts,
- one daily-state recompute and one statistics rewrite.

Files:

- `Python/rehab_excel/daily_replacement_fact_batch.py`
- `Python/tools/restore_daily_replacement_facts.py`
- `Tests/test_daily_replacement_fact_batch.py`
- `Tests/test_authoritative_bulk_daily_restore.py`

## Current validation state

The focused source gate is green:

- **36 passed in 4.74s**

The first V89 -> V90 native re-placement attempt then exposed a separate persistence bug:

`SAFETY STOP: Duplicate ReplacementId 'daily-input:planner:61:ΦΘ@2026-10-19:2026-10-19'`

Cause:

- existing ACTIVE facts were correctly loaded for operational reconstruction,
- but `apply_daily_input_replacements.py` attempted to append the complete active set again after adding the new re-placement,
- the duplicate-id safety guard stopped the write before a duplicate could be persisted.

Fix implemented:

- only replacement assignments created during the current run are appended as new durable facts,
- previously persisted ACTIVE facts are used for state/workload only and are not re-appended,
- if a failure occurs after the preview copy is created but before persistence completes, the newly created preview is deleted to avoid leaving a partial workbook,
- reporting now distinguishes loaded, newly accepted, and total active replacements.

Commits:

- `ae0f58d363ec0cbf357876cf6b4633e3f14d70e7` — rerun persistence + partial-preview cleanup
- `88b75319273005b28d373d2a683701f82fe41fba` — regression coverage

Current status: **fix committed; Windows focused gate rerun required before recreating V90**.

Any `RERESOLVED_V90.xlsm` from the failed run is untrusted/partial and must be overwritten from V89 after the gate passes.

## Immediate roadmap

1. Pull latest branch and rerun the focused source gate.
2. Recreate V90 from V89 with `apply_daily_input_replacements.py`; choose a candidate and require new fact id ending `:v2`.
3. Inspect V90 and confirm old session/date fact remains `RESTORED` while new `:v2` fact is `ACTIVE`.
4. Verify resulting V90 reconstructs with 5 replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ and exact statistics equivalence.
5. Test selected authoritative restore of two ACTIVE facts into V91; expected 2 restored, 3 active, unresolved=2, finalized=ΟΧΙ.
6. Test `--all-active` restore from V90 into V92; expected 5 restored, 0 active, unresolved=5, finalized=ΟΧΙ.
7. Connect Excel ledger per-row/selected/all controls to authoritative backend; retire old cell-copy VBA only after native bridge gate.
8. Resume outpatient recurring create/edit native gate.
9. Later clean Python 3.14 ZipFile warnings.

## Locked daily rules

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. unresolved orphans resolve through placement or treatment suspension,
7. only then expose capacity to new patients.

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Documentation discipline

Every meaningful evolution must be recorded in the repository documentation during the same work cycle, including implemented changes, regressions, native validation results, newly discovered necessities, planned additions, removals, replacements and deferred fixes. Chat history is not the sole source of truth.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Safety constraints

- never write directly to baseline `.xlsm` during development,
- preview workflows use copies,
- preserve and verify VBA,
- keep layout write zones narrow,
- existing replacement facts must never be re-appended during a rerun,
- partial preview files created before a later persistence failure must be removed,
- bulk restore validates the whole set before any mutation,
- Windows tests use `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
