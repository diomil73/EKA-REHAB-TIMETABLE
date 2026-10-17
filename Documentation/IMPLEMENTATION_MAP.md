# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Ορφανοί: operational queue με τελικό outcome

Αυτό είναι πλέον το ενεργό milestone.

Κάθε unresolved/displaced patient για τη συγκεκριμένη ημερομηνία πρέπει να καταλήγει υποχρεωτικά σε ένα από:

- `Τοποθετήθηκε` σε έγκυρο therapist/time,
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Locked rules:

- κανένας ορφανός δεν εξαφανίζεται σιωπηλά,
- `Αναστολή θεραπείας` είναι daily overlay/status, όχι recurring edit,
- πρέπει να είναι auditable,
- πρέπει να τροφοδοτεί daily statistics ως μη πραγματοποιηθείσα θεραπεία με explicit reason,
- displaced existing patients προηγούνται των νέων ασθενών,
- recomputation παραμένει idempotent και order-independent.

### 2. Daily effective state: MASTER + THERAPIST_DAILY + ledger

Το stress case ΜΕΝΤΗΣ είναι native green για:

- active move `Σκαρώνης 09:15 -> Μηλιδάκης 10:45`,
- MASTER + THERAPIST_DAILY sync,
- source trace,
- individual restore,
- readable `hh:mm` time axis,
- stale robotic-orange leak fix.

Statistics integration παραμένει open και συνδέεται τώρα με το orphan/suspension milestone.

### 3. Legacy `ΑΝΑΝΕΩΣΗ`

Το κουμπί παραμένει οπτικά, αλλά είναι **deferred/manual cleanup**.

Έχει ταυτοποιηθεί ως Form Control `Button 1`, caption `ΑΝΑΝΕΩΣΗ`, macro `RefreshAll`. Δεν είναι λειτουργικό blocker και δεν θα καταναλώσουμε άλλο engineering time τώρα. Ο χρήστης μπορεί να το αφαιρέσει χειροκίνητα στο Excel αργότερα.

### 4. Reversible daily change ledger

Native green:

- per-row `Επαναφορά`,
- MASTER sync,
- source trace,
- styling cleanup,
- time-axis display.

Open:

- selected/all restore σε multi-change preview,
- production restore = authoritative overlay removal + recompute, όχι snapshot restore.

### 5. Dynamic daily recomputation

Locked order:

1. recurring/base schedule για concrete date,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute στα net free slots,
6. resolve remaining orphans με placement ή `Αναστολή θεραπείας`,
7. μόνο μετά new-patient allocation.

### 6. Replacement operational load semantics

- authoritative load = occupied timeslots,
- shared slot με πολλούς patients = ένα occupied timeslot,
- ranking πρώτα occupied timeslots, μετά session count,
- output `load x/y | patients n`.

### 7. Outpatient recurring create/edit

Παραμένει open και επιστρέφει αμέσως μετά το current orphan/daily-state milestone.

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
- V82 individual restore: native green.
- V83 MASTER sync + restore: native green.
- V84 time-axis + restore: native green.
- V85 stale-orange cleanup + restore: native green.

## Τρέχον verification gate

1. Inspect daily-state, cancellation, reason/status and statistics models.
2. Choose clean representation for date-specific `Αναστολή θεραπείας`.
3. Implement orphan placement/suspension outcomes.
4. Add statistics integration.
5. Add focused regression tests for auditability, non-destructive recurring behavior and idempotent recompute.
6. Native-check orphan workflow.
7. Test selected/all restore with multiple changes.
8. Replace snapshot restore with authoritative overlay removal + recompute.
9. Resume outpatient recurring create/edit gate.
10. Clean Python 3.14 ZipFile teardown warnings later.

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

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
