# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Authoritative daily facts / restart-safe recomputation

Το mixed orphan workflow είναι native green και οι αναστολές επιβιώνουν σε fresh reread.

V87 confirmed:

- 5 placements,
- 2 `Αναστολή θεραπείας`,
- unresolved=0,
- scheduled=80,
- performed=78,
- replaced=5,
- suspended=2,
- source unchanged=True,
- VBA preserved=True.

V87 persistence verifier confirmed:

- 2 suspensions reread from DAILY_INPUT,
- stored suspended=2,
- stored unresolved=0,
- stored finalized=ΝΑΙ.

Durable replacement placement persistence is implemented and its focused Windows source gate is green.

Implemented slice:

- `rehab_excel/daily_replacement_facts.py`,
- internal `DAILY_REPLACEMENT_FACTS` very-hidden sheet,
- exact structured ReplacementAssignment fields,
- validation against concrete Session,
- duplicate/conflict rejection,
- idempotent persistence,
- backward-compatible missing sheet = no persisted replacements,
- deterministic replacement ids per session/date,
- `apply_daily_input_replacements.py` reloads persisted replacements, uses them in capacity/state, skips already-resolved sessions and persists new placements,
- `verify_daily_fact_reconstruction.py` reconstructs daily state from base + DAILY_INPUT + DAILY_REPLACEMENT_FACTS and compares with DAILY_STATISTICS.

Focused Windows gate: **28 passed, 3 warnings in 5.14s**.

V88 exposed one native reader issue: `ReplacementDate` was persisted by COM as timezone-shifted numeric serial `46313.875`. The fix now:

- normalizes legacy numeric calendar dates to nearest whole Excel day,
- explicitly covers `46313.875 -> 2026-10-19`,
- writes future ReplacementDate facts as ISO text `YYYY-MM-DD` to avoid timezone ambiguity.

Status: **V88 date-reader fix implemented; focused regression + reconstruction rerun pending**.

### 2. Ορφανοί + treatment suspension

Locked rule: κάθε orphan πρέπει να καταλήγει σε:

- valid therapist/time placement, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`.
Recurring/base δεν αλλάζει. Finalized day μόνο με unresolved=0.

Operational flow source + native green through V87.

### 3. Daily statistics

`DAILY_STATISTICS` derives from the same finalized DailySessionState and stores scheduled, performed, replaced, patient absent, department postponed, no-show, treatment suspended, unresolved and finalized.

Native-created in V87. Full reconstruction equivalence remains the V88 gate.

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

- authoritative restore by removal/deactivation of daily facts + recompute,
- selected/all restore in a multi-change preview.

Replacement fact persistence is the prerequisite foundation for this.

### 6. Dynamic daily recomputation

Locked order:

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced patients,
5. load/apply persisted replacement overlays and redistribute existing patients,
6. remaining orphans resolve by placement or suspension,
7. new patients only after existing displaced patients are resolved.

Recompute must be idempotent and order-independent.

### 7. Outpatient recurring create/edit

Implemented but native verification remains open after authoritative daily facts / restore milestone.

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
- V87 suspension/statistics reread: **native green**.
- Replacement fact + reconstruction focused source gate: **28 passed, 3 warnings in 5.14s**.

## Τρέχον verification gate

1. Pull latest branch.
2. Run focused `Tests/test_daily_replacement_facts.py` regression.
3. Re-run reconstruction directly against existing V88. No rebuild required for backward-compatible reader verification.
4. Require 5 replacements, 2 suspensions, unresolved=0, finalized=ΝΑΙ and exact DAILY_STATISTICS equivalence.
5. If green, close authoritative daily-state restart persistence.
6. Implement ledger restore via authoritative fact removal + recompute.
7. Test selected/all restore with multiple changes.
8. Resume outpatient recurring create/edit native gate.
9. Clean Python 3.14 ZipFile teardown warnings later.

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
