# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Authoritative restore + recompute

Το restart-safe daily truth είναι native green στο V88 και το πρώτο πραγματικό authoritative restore είναι native green στο V89.

V89 confirmed:

- exact replacement fact became `RESTORED`,
- other 4 facts stayed `ACTIVE`,
- reread replacements=4,
- suspensions=2,
- unresolved=1,
- finalized=ΟΧΙ,
- reconstructed statistics = stored DAILY_STATISTICS.

Άρα το restore πλέον σημαίνει fact lifecycle + recompute, όχι cell copy-back.

### 2. Replacement ID hardening

Implemented:

- first placement keeps legacy id `daily-input:<session_id>:<date>`,
- re-placement after restore gets `:v2`,
- later cycles get `:v3`, etc.,
- full replacement history including RESTORED rows drives version allocation.

This prevents an old restore request from targeting a later change.

### 3. Selected/all authoritative restore

Implemented backend:

- `daily_replacement_fact_batch.py` validates the whole requested set before mutation,
- missing/stale/already-restored member aborts the complete batch,
- no partial restore,
- selected restore accepts multiple exact ids,
- restore-all selects all currently ACTIVE facts,
- one batch mutation + one recompute + one statistics rewrite.

Tool:

- `Python/tools/restore_daily_replacement_facts.py`

Status: **source/native gate pending**.

### 4. Ορφανοί + treatment suspension

Native green through V89 reconstruction.

Locked rule: κάθε orphan καταλήγει σε valid therapist/time placement ή `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα. Recurring/base δεν αλλάζει. Finalized day μόνο με unresolved=0.

### 5. Daily statistics

`DAILY_STATISTICS` derives from the same DailySessionState. V88 and V89 both prove fresh reconstruction gives exact stored-statistics equivalence.

### 6. Daily effective state: MASTER + THERAPIST_DAILY

Native green for active replacements. Authoritative restore resets MASTER to recurring truth before reapplying surviving overlays and rebuilds THERAPIST_DAILY from recomputed state.

### 7. Excel ledger migration

Earlier prototype native green for per-row UI, source trace, readable times and styling.

Still open:

- connect per-row button to authoritative Python bridge,
- connect selected/all buttons to atomic authoritative batch restore,
- retire old snapshot cell-copy VBA only after native bridge gate.

### 8. Dynamic daily recomputation

Locked order:

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply ACTIVE persisted replacements,
6. remaining orphans resolve by placement or suspension,
7. new patients only after existing displaced patients are resolved.

Recompute remains idempotent and order-independent.

### 9. Outpatient recurring create/edit

Implemented but native verification remains open after ledger authoritative-restore migration.

### 10. Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cleanup. Not a blocker.

## Ολοκληρωμένα checkpoints

- full regression: **699 passed, 19 warnings**
- load + mixed-slot: **31 passed in 2.04s**
- MASTER overlay + restore: **24 passed, 5 warnings in 4.39s**
- suspension/statistics domain: **21 passed in 3.11s**
- orphan operational: **43 passed, 3 warnings in 5.15s**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- V87 mixed orphan flow: **native green**
- V88 full restart reconstruction: **native green**
- V89 single authoritative restore + reconstruction: **native green**

Python 3.14 ZipFile teardown warnings remain non-blocking technical debt.

## Τρέχον verification gate

1. Pull latest branch.
2. Run version-id + batch restore + authoritative restore focused tests.
3. Re-resolve the reopened V89 orphan and require a new `:v2` replacement fact.
4. Verify fresh reconstruction returns unresolved=0/finalized=ΝΑΙ and exact statistics equivalence.
5. Test selected restore of two ACTIVE facts.
6. Test restore-all on a separate preview and verify no partial corruption.
7. Wire Excel ledger buttons to authoritative backend.
8. Resume outpatient create/edit native gate.

## Business rules που δεν πρέπει να παλινδρομήσουν

- daily facts are overlays, never destructive recurring edits,
- operational capacity is timeslot-based,
- existing displaced patients precede new patients,
- orphan never silently disappears,
- MASTER and THERAPIST_DAILY agree on effective state,
- old restore cannot overwrite a newer change,
- bulk restore is validate-all-before-mutate,
- layout write zones remain narrow,
- ΦΘ provider remains solver-owned,
- outpatients cannot receive robotic treatment.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
