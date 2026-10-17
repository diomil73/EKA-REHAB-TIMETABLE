# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Ορφανοί: operational queue + date-specific suspension

Το mixed native scenario είναι πλέον green.

Κάθε orphan πρέπει να καταλήγει υποχρεωτικά σε:

- έγκυρη τοποθέτηση therapist/time, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Architecture locked:

- suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`,
- recurring/base δεν αλλάζει,
- original slot απελευθερώνεται,
- reason auditable,
- finalized day μόνο με unresolved=0.

Source status:

- domain foundation green,
- mandatory placement/suspension flow green,
- suspension persistence green,
- DAILY_INPUT reread support green,
- suspension dropdown green,
- DAILY_STATISTICS writer green,
- consolidated operational gate: **43 passed, 3 warnings in 5.15s**.

Native V87 status:

- 5 accepted replacements,
- 2 treatment suspensions,
- 2 persisted suspension facts,
- unresolved=0,
- scheduled=80,
- performed=78,
- replaced=5,
- suspended=2,
- source unchanged=True,
- VBA preserved=True.

### 2. Restart persistence / authoritative daily facts

Current truth split:

- `Αναστολή θεραπείας` is durably stored in output DAILY_INPUT and can be reread.
- replacement placements are not yet durable authoritative reread facts.

New verifier:

- `Python/tools/verify_orphan_resolution_persistence.py`
- regression: `Tests/test_verify_orphan_resolution_persistence.py`

It verifies suspension reread and stored DAILY_STATISTICS counts without pretending that placements already survive restart.

Next implementation target: persist replacement assignments as authoritative date-specific facts, then rebuild the full daily effective state from those facts after restart.

### 3. Daily statistics

`DAILY_STATISTICS` is native-created in V87 and derives from the same finalized `DailySessionState` used during resolution.

Fields: scheduled, performed, replaced, patient absent, department postponed, no-show, treatment suspended, unresolved, finalized.

### 4. Daily effective state: MASTER + THERAPIST_DAILY + ledger

MENTIS stress case remains native green for active move, MASTER/THERAPIST_DAILY sync, source trace, individual restore, readable times and stale robotic-orange fix.

### 5. Reversible daily change ledger

Native green:

- per-row restore,
- MASTER sync,
- source trace,
- styling cleanup,
- time-axis display.

Open:

- selected/all restore in a multi-change preview,
- production restore through authoritative overlay removal + full recompute.

### 6. Dynamic daily recomputation

Locked order:

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute existing patients,
6. resolve remaining orphans with placement or suspension,
7. only then allocate new patients.

Recompute idempotent and order-independent.

### 7. Outpatient recurring create/edit

Implemented but native verification remains open after the authoritative daily-fact milestone.

### 8. Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cleanup. Not a blocker.

## Τι έχει ολοκληρωθεί

- MASTER operational home + accepted V81 toolbar/header UX.
- DAILY_INPUT authoritative input flow.
- Patient-centric multi-treatment scheduling.
- Edit-existing-program preload + atomic resave.
- Conflict validation + explicit override.
- Optional manual provider/name persistence.
- Retired single-treatment cleanup.
- Full regression checkpoint: **699 passed, 19 warnings**.
- Cancellation slot release support.
- Replacement load + mixed-slot gate: **31 passed in 2.04s**.
- MASTER daily overlay + restore gate: **24 passed, 5 warnings in 4.39s**.
- V82 restore native green.
- V83 MASTER sync native green.
- V84 time-axis native green.
- V85 stale-orange cleanup native green.
- Suspension/statistics domain gate: **21 passed in 3.11s**.
- Consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**.
- V87 mixed orphan resolution: **native green**.

## Τρέχον verification gate

1. Run persistence verifier against V87: suspended=2, unresolved=0, finalized=ΝΑΙ.
2. Persist replacement placements as authoritative date-specific facts.
3. Add reread/recompute test proving 5 replacements + 2 suspensions reconstruct the same daily state.
4. Connect restore to authoritative fact removal + recompute.
5. Test selected/all restore with multiple changes.
6. Resume outpatient recurring create/edit native gate.
7. Clean Python 3.14 ZipFile teardown warnings later.

## Business rules που δεν πρέπει να παλινδρομήσουν

- ΦΘ provider required + solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, όχι solver constraint.
- Outpatients cannot receive Ρομποτικό.
- Ανακλινόμενο / Πισίνα / ΕΦΑ χωρίς provider/name.
- Nonblank manual constraints δεν αντικαθίστανται σιωπηλά.
- Standard duration 45', Ανακλινόμενο 30'.
- Exact adjacency allowed.
- `PATIENT_PLANNER` authoritative για inpatient recurring, `OUTPATIENT_SCHEDULE` για outpatient recurring.
- Daily facts είναι overlays, όχι recurring edits.
- Existing displaced patients προηγούνται των νέων.
- Orphans πρέπει να τοποθετηθούν ή να πάρουν explicit suspension.
- MASTER και THERAPIST_DAILY συμφωνούν στο effective daily state.
- Operational capacity είναι timeslot-based.
- Layout write zones παραμένουν στενές.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
