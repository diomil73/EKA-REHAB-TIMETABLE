# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **authoritative restore by daily-fact lifecycle + recompute**.

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics domain gate: **21 passed in 3.11s**
- consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- V87 mixed orphan scenario: **native green**
- V87 suspension/statistics reread: **native green**
- V88 full daily-state reconstruction across restart: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

Fresh verification against `ORPHAN_RESOLUTION_V88.xlsm` confirmed:

- persisted replacements reread: 5
- treatment suspensions reread: 2
- reconstructed unresolved: 0
- reconstructed finalized: ΝΑΙ
- stored DAILY_STATISTICS exactly matches reconstructed DailySessionState

The operational day now survives restart through durable facts instead of relying on rendered workbook cells.

## ACTIVE: authoritative restore

The current Excel ledger macro still contains the old snapshot-copy implementation. New source foundation now exists to replace it safely.

`daily_replacement_facts.py` now supports `ACTIVE` and `RESTORED` lifecycle:

- RESTORED remains audit history,
- RESTORED is excluded from effective daily state,
- exact restore uses immutable ReplacementId,
- stale undo cannot silently overwrite a newer fact,
- same historical ReplacementId cannot be reused.

New restore preview tool:

- `Python/tools/restore_daily_replacement_fact.py`

It performs:

1. exact fact ACTIVE -> RESTORED,
2. fresh daily-state reconstruction,
3. statistics recomputation,
4. MASTER recurring projection reset,
5. THERAPIST_DAILY rebuild,
6. surviving MASTER overlay reapplication,
7. DAILY_STATISTICS rewrite,
8. ledger history mark RESTORED,
9. source/VBA preservation verification.

New regression:

- `Tests/test_authoritative_daily_restore.py`

Inventory helper:

- `Python/tools/inspect_daily_replacement_facts.py`

Status: **implemented; focused source gate pending; first native authoritative restore V89 pending**.

## Important restore behavior

Undoing a placement does not imply the original session can run. If its therapist absence remains active, the session becomes orphan/unresolved again.

Therefore after restoring one of the five V88 placements, the expected truth is likely:

- active replacements: 4
- treatment suspensions: 2
- unresolved: 1
- finalized: ΟΧΙ

That is correct until the reopened orphan is re-placed or explicitly suspended.

## Next hardening: replacement IDs

Current IDs are `daily-input:<session_id>:<date>`. Once that ID is RESTORED, a later new placement for the same session/date must receive a new/versioned ID. The writer intentionally rejects historical ID reuse. Versioned ID generation follows the first successful V89 native restore.

## Immediate roadmap

1. Run focused lifecycle + restore tests.
2. Inventory ACTIVE facts in V88.
3. Restore one exact replacement into `AUTHORITATIVE_RESTORE_V89.xlsm`.
4. Run full reconstruction verifier against V89 and verify recomputed statistics.
5. Connect ledger per-row button to authoritative Python bridge instead of snapshot copy-back.
6. Implement selected/all restore as safe multi-fact lifecycle change + one recompute.
7. Add replacement-ID versioning for re-resolution.
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

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Safety constraints

- never write directly to baseline `.xlsm` during development,
- preview workflows use copies,
- preserve and verify VBA,
- keep layout write zones narrow,
- Windows tests use `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
