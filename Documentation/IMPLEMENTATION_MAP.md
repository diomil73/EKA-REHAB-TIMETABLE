# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Authoritative restore + recompute

Το restart-safe daily truth είναι πλέον **native green** στο V88:

- 5 persisted replacements reread,
- 2 treatment suspensions reread,
- unresolved=0,
- finalized=ΝΑΙ,
- reconstructed statistics = stored DAILY_STATISTICS.

Άρα κλείνει το milestone:

`base + DAILY_INPUT + DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Τώρα αντικαθιστούμε το προσωρινό snapshot undo του ledger με authoritative fact lifecycle.

Implemented:

- `DAILY_REPLACEMENT_FACTS` statuses `ACTIVE / RESTORED`,
- RESTORED rows παραμένουν audit history αλλά δεν συμμετέχουν στο effective state,
- exact restore by immutable ReplacementId,
- stale/older undo δεν μπορεί να σβήσει νεότερη αλλαγή,
- νέο `Python/tools/restore_daily_replacement_fact.py`,
- MASTER reset σε recurring truth πριν από reapplication surviving overlays,
- THERAPIST_DAILY rebuild από recomputed DailySessionState,
- DAILY_STATISTICS rewrite,
- ledger history mark RESTORED,
- source/VBA preservation checks,
- `inspect_daily_replacement_facts.py` για exact native fact inventory.

Status: **implemented, focused source gate pending, native V89 pending**.

Σημαντική semantics: αν γίνει restore ενός placement ενώ η therapist absence εξακολουθεί να ισχύει, ο ασθενής ξαναγίνεται orphan. Άρα είναι σωστό να προκύψει unresolved>0 και finalized=ΟΧΙ μέχρι νέα τοποθέτηση ή `Αναστολή θεραπείας`.

### 2. Replacement ID hardening

Σήμερα τα operational IDs είναι deterministic:

`daily-input:<session_id>:<date>`

Μετά από RESTORED fact δεν πρέπει να ξαναχρησιμοποιηθεί το ίδιο historical ID. Ο writer πλέον το απορρίπτει. Μετά το πρώτο native authoritative restore gate θα προστεθεί versioned/new ID generation για re-resolution της ίδιας session/date.

### 3. Ορφανοί + treatment suspension

Native green through V88 persistence/reconstruction.

Locked rule: κάθε orphan καταλήγει σε:

- valid therapist/time placement, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`.
Recurring/base δεν αλλάζει. Finalized day μόνο με unresolved=0.

### 4. Daily statistics

`DAILY_STATISTICS` derives from the same DailySessionState. V88 proves fresh reconstruction gives exact stored-statistics equivalence.

### 5. Daily effective state: MASTER + THERAPIST_DAILY

Native green for active replacement sync. Authoritative restore pipeline now explicitly resets MASTER to recurring projection and re-applies only surviving ACTIVE overlays.

### 6. Reversible daily change ledger

Earlier prototype native green:

- per-row snapshot restore,
- MASTER sync,
- source trace,
- styling cleanup,
- readable time axis.

Current migration:

- authoritative per-row restore via fact RESTORED + recompute,
- then selected/all restore by multi-fact change + one recompute.

### 7. Dynamic daily recomputation

Locked order:

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply ACTIVE persisted replacements,
6. remaining orphans resolve by placement or suspension,
7. new patients only after existing displaced patients are resolved.

Recompute idempotent and order-independent.

### 8. Outpatient recurring create/edit

Implemented but native verification remains open after authoritative restore milestone.

### 9. Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cleanup. Not a blocker.

## Ολοκληρωμένα checkpoints

- full regression: **699 passed, 19 warnings**
- load + mixed-slot: **31 passed in 2.04s**
- MASTER overlay + restore: **24 passed, 5 warnings in 4.39s**
- suspension/statistics domain: **21 passed in 3.11s**
- orphan operational: **43 passed, 3 warnings in 5.15s**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- V87 mixed orphan flow: **native green**
- V87 suspension persistence: **native green**
- V88 full restart reconstruction: **native green**

Python 3.14 ZipFile teardown warnings remain non-blocking technical debt.

## Τρέχον verification gate

1. Run focused tests for fact lifecycle + authoritative restore pipeline + reconstruction verifier.
2. Inventory V88 ACTIVE replacement IDs.
3. Restore one exact fact into V89.
4. Expected after one restore: active replacements 4; if absence still exists, unresolved 1 and finalized=ΟΧΙ.
5. Re-run reconstruction verifier against V89 and require exact stored statistics equivalence.
6. Connect Excel ledger per-row button to authoritative bridge.
7. Implement selected/all authoritative restore and one recompute.
8. Add replacement-ID versioning for re-resolution.
9. Resume outpatient create/edit native gate.

## Business rules που δεν πρέπει να παλινδρομήσουν

- daily facts are overlays, never destructive recurring edits,
- operational capacity is timeslot-based,
- existing displaced patients precede new patients,
- orphan never silently disappears,
- MASTER and THERAPIST_DAILY agree on effective state,
- old restore cannot overwrite a newer change,
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
