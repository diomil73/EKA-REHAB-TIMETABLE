# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest native daily preview: `Excel/previews/ORPHAN_RESOLUTION_V88.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement load + mixed outpatient/inpatient: `31 passed in 2.04s`
- MASTER daily overlay + ledger restore source gate: `24 passed, 5 warnings in 4.39s`
- orphan suspension + statistics domain gate: `21 passed in 3.11s`
- consolidated orphan operational gate: `43 passed, 3 warnings in 5.15s`
- replacement fact focused gate: `28 passed, 3 warnings in 5.14s`
- V87 mixed orphan resolution: native green
- V87 suspension/statistics reread: native green
- V88 full daily-fact reconstruction across restart: **native green**

Known Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED milestone: restart-safe daily truth

`ORPHAN_RESOLUTION_V88.xlsm` fresh-process reconstruction confirmed:

- persisted replacements reread: 5
- treatment suspensions reread: 2
- reconstructed unresolved: 0
- reconstructed finalized: ΝΑΙ
- stored `DAILY_STATISTICS` exactly matches reconstructed daily state

Therefore the concrete operational day can now be reconstructed from durable facts:

`recurring/base + DAILY_INPUT + DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

The earlier COM date issue (`46313.875`) is also closed:

- legacy numeric ReplacementDate values normalize to nearest whole Excel day,
- future writes use ISO text `YYYY-MM-DD`.

## ACTIVE milestone: authoritative restore + recompute

The old VBA `RestoreLedgerRow()` still performs temporary snapshot cell copy-back. We are replacing that architecture with fact lifecycle + recomputation.

### Replacement fact lifecycle implemented

`Python/rehab_excel/daily_replacement_facts.py` now supports:

- `ACTIVE`
- `RESTORED`

Rules:

- RESTORED rows remain immutable audit history,
- effective daily-state readers ignore RESTORED facts,
- conflict checks only consider ACTIVE facts,
- exact restore matches immutable `ReplacementId`, not cells,
- already-restored fact is idempotent,
- a stale restore cannot silently erase a newer replacement,
- unsupported statuses safety-stop.

Function:

- `restore_daily_replacement_fact(workbook_path, replacement_id)`

Commits:

- lifecycle implementation: `d5bba940e6ca531fd5a969409941aeb6def5b8e3`
- lifecycle regressions: `14c0566c04c9489fe7ec1a88a74779e2b68d0371`

### Authoritative restore preview implemented

New tool:

- `Python/tools/restore_daily_replacement_fact.py`

Flow:

1. copy source workbook to staging,
2. mark exact replacement fact RESTORED,
3. reread all surviving facts,
4. rebuild DailySessionState,
5. recompute statistics,
6. rebuild MASTER back to recurring truth with `refresh_master_projection_in_place`,
7. rebuild THERAPIST_DAILY from recomputed state,
8. reapply surviving MASTER daily overlays,
9. rewrite DAILY_STATISTICS,
10. mark matching ledger ChangeId as RESTORED history,
11. preserve source and VBA.

Commit: `dfc7c0424361eb293110e79f34545f15792bf20d`

Regression:

- `Tests/test_authoritative_daily_restore.py`
- commit `29c28397c478a17ab69d24d1cba0e52c27128270`

Fact inventory helper:

- `Python/tools/inspect_daily_replacement_facts.py`
- commit `3d622d4003e17252af58ce0dcd9b1364797b1c36`

Status: **implemented; focused source gate pending; native V89 restore pending**.

## Important restore semantic

Restoring a replacement means removing that daily placement overlay. If the original therapist absence still exists, that patient becomes orphan/unresolved again.

That is correct behavior. The system must not silently put the patient back under an absent therapist.

Therefore a valid authoritative restore can produce:

- fewer active replacements,
- `unresolved > 0`,
- `finalized = ΟΧΙ`,

until that orphan is placed again or receives `Αναστολή θεραπείας`.

## ID hardening item

Current operational flow uses deterministic replacement IDs:

`daily-input:<session_id>:<date>`

After one such fact becomes RESTORED, a later new replacement for the same session/date must not reuse that historical ID. The writer now rejects reuse intentionally.

Next hardening after the first native restore gate: add versioned/new replacement-ID generation for re-resolution of a previously restored session.

## Exact resume point

1. Run focused lifecycle + authoritative restore tests.
2. Inventory exact ACTIVE replacement IDs in V88.
3. Restore one exact V88 replacement into `AUTHORITATIVE_RESTORE_V89.xlsm`.
4. Expect active replacements to drop from 5 to 4. If therapist absence persists, expect unresolved to rise to 1 and finalized to become ΟΧΙ.
5. Run fresh reconstruction verifier against V89 and require exact DAILY_STATISTICS equivalence.
6. After V89 is native green, connect Excel ledger buttons to this authoritative path through an existing bridge pattern.
7. Implement selected/all restore as multi-fact deactivation followed by one recompute.
8. Add replacement-ID versioning for re-resolution after restore.
9. Resume outpatient recurring create/edit native gate.

## Locked orphan rule

Every orphan for the concrete date must resolve as exactly one of:

- valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`.
Recurring/base remains untouched. A finalized day requires unresolved=0.

## Locked daily recomputation order

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacement overlays,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose capacity to new patients.

Operational capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree on effective daily state.

## Reversible daily change ledger

Native green from earlier prototype:

- active move rendering,
- MASTER sync,
- source trace,
- per-row snapshot restore,
- readable time axis,
- stale-orange cleanup.

Open:

- replace snapshot restore with authoritative fact restore + recompute,
- selected/all authoritative restore with multiple simultaneous changes.

## Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cosmetic cleanup. Not a blocker. Do not reopen unless operationally relevant.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
