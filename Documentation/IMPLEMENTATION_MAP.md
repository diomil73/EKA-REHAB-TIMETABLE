# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Ορφανοί: operational queue + date-specific suspension

Ενεργό milestone.

Κάθε unresolved/displaced patient για τη συγκεκριμένη ημερομηνία πρέπει να καταλήγει υποχρεωτικά σε:

- `Τοποθετήθηκε` σε έγκυρο therapist/time, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Architecture locked:

- suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`,
- recurring/base δεν αλλάζει,
- original slot απελευθερώνεται,
- reason παραμένει auditable,
- ημέρα finalized μόνο με `unresolved_sessions == 0`.

Source status:

- domain foundation green,
- mandatory placement/suspension flow green,
- suspension persistence green,
- DAILY_INPUT reread support green,
- `Αναστολή θεραπείας` dropdown green,
- DAILY_STATISTICS writer green,
- consolidated operational gate: **43 passed, 3 warnings in 5.15s**.

### 2. Native mixed-orphan verification

Το πρώτο V87 attempt σταμάτησε σωστά από layout safety πριν γραφτεί workbook:

`THERAPIST_DAILY!A21 outside confirmed write zones B2:J8, B12:J18`.

Root cause: stale outpatient-fill cleanup έκανε scan σε όλο το sheet και θεώρησε unrelated πράσινο κελί `A21` ως outpatient presentation residue.

Fix implemented:

- stale outpatient cleanup περιορίζεται μόνο στα confirmed operational grids `B2:J8` και `B12:J18`,
- οι write zones **δεν διευρύνθηκαν**,
- `A21`/ledger/labels/other furniture αγνοούνται,
- regression καλύπτει explicit το `A21`.

Commits:

- `5c81ce40256ddd36379d838c35b2a62debf0ef95`
- `7e4031fd9f01b4b11233b6bc087078b4f94686aa`

Native smoke target παραμένει:

- input `Excel/previews/ORPHAN_SMOKE_INPUT_V86.xlsm`,
- output `Excel/previews/ORPHAN_RESOLUTION_V87.xlsm`,
- τουλάχιστον μία placement και μία `Αναστολή θεραπείας`.

### 3. Daily statistics

Το `DAILY_STATISTICS` sheet γράφεται από το ίδιο finalized `DailySessionState` και περιλαμβάνει scheduled, performed, replaced, patient absent, department postponed, patient no-show, treatment suspended, unresolved και finalized.

Source green. Native workbook verification pending.

### 4. Daily effective state: MASTER + THERAPIST_DAILY + ledger

Το stress case ΜΕΝΤΗΣ είναι native green για active move, MASTER + THERAPIST_DAILY sync, source trace, individual restore, readable time axis και stale robotic-orange fix.

### 5. Legacy `ΑΝΑΝΕΩΣΗ`

**Deferred/manual cleanup**. Δεν είναι blocker.

### 6. Reversible daily change ledger

Native green: per-row restore, MASTER sync, source trace, styling cleanup, time-axis display.

Open: selected/all restore σε multi-change preview και production restore μέσω authoritative overlay removal + recompute.

### 7. Dynamic daily recomputation

Locked order:

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute στα net free slots,
6. resolve remaining orphans με placement ή `Αναστολή θεραπείας`,
7. μόνο μετά new-patient allocation.

Recompute idempotent και order-independent.

### 8. Replacement operational load semantics

Authoritative load = occupied timeslots. Shared slot με πολλούς patients = ένα occupied timeslot. Ranking πρώτα occupied timeslots, μετά session count.

### 9. Outpatient recurring create/edit

Παραμένει open μετά το current orphan/daily-state milestone.

## Τι έχει ολοκληρωθεί

- MASTER operational home + accepted V81 toolbar/header UX.
- DAILY_INPUT authoritative flow.
- Patient-centric multi-treatment scheduling.
- Edit-existing-program preload + atomic resave.
- Conflict validation + explicit override.
- Optional manual provider/name persistence.
- Retired single-treatment cleanup.
- Full regression checkpoint: **699 passed, 19 warnings**.
- Cancellation slot release support.
- Replacement load + mixed-slot gate: **31 passed in 2.04s**.
- MASTER daily overlay + restore gate: **24 passed, 5 warnings in 4.39s**.
- V82 individual restore native green.
- V83 MASTER sync + restore native green.
- V84 time-axis + restore native green.
- V85 stale-orange cleanup + restore native green.
- Suspension/statistics domain slice: **21 passed in 3.11s**.
- Consolidated orphan operational flow/persistence/statistics gate: **43 passed, 3 warnings in 5.15s**.

## Τρέχον verification gate

1. Run focused outpatient-presentation + orphan tests after A21 fix.
2. Re-run V87 from existing V86.
3. Resolve one orphan by placement and one by `S = Αναστολή θεραπείας`.
4. Verify DAILY_INPUT persisted fact + DAILY_STATISTICS unresolved=0/finalized=ΝΑΙ + source/VBA safety.
5. Re-read output to confirm persistence.
6. Test selected/all restore with multiple changes.
7. Replace snapshot restore with authoritative overlay removal + recompute.
8. Resume outpatient recurring create/edit gate.
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
- Daily cancellations/absences/replacements/suspensions/redistributions είναι overlays, όχι recurring edits.
- Existing displaced patients προηγούνται των νέων ασθενών.
- Orphans πρέπει να τοποθετηθούν ή να πάρουν explicit `Αναστολή θεραπείας`.
- MASTER και THERAPIST_DAILY πρέπει να συμφωνούν στο effective daily state.
- Operational capacity είναι timeslot-based.
- Layout safety zones παραμένουν στενές. Cleanup logic προσαρμόζεται στο contract, όχι το αντίστροφο.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
