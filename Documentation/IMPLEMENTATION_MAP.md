# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Ενοποίηση daily effective state σε MASTER + THERAPIST_DAILY

Το V82 απέδειξε ότι το `THERAPIST_DAILY` ακολουθεί σωστά το active daily replacement, αλλά το timetable στο `MASTER` μένει στο recurring/base πρόγραμμα. Αυτό είναι πλέον το πρώτο integration bug προς διόρθωση.

Locked target:

`DailySessionState -> THERAPIST_DAILY + MASTER timetable + ledger + statistics`

Native acceptance για το stress case ΜΕΝΤΗΣ:

- active overlay: `Σκαρώνης 09:15 -> Μηλιδάκης 10:45`,
- MASTER πρέπει να δείχνει `Μηλιδάκης 10:45` όσο η αλλαγή είναι ενεργή,
- μετά `Επαναφορά`, MASTER πρέπει να ξαναδείξει `Σκαρώνης 09:15`,
- κανένα view δεν πρέπει να κρατά δεύτερη/παλιά operational αλήθεια.

### 2. V82 reversible daily change ledger

Το πραγματικό stress test είναι native λειτουργικό:

- date: `2026-10-19`
- patient: `ΜΕΝΤΗΣ`
- original: `Σκαρώνης 09:15`
- replacement: `Μηλιδάκης 10:45`
- Μηλιδάκης: `load 5/6 | patients 6`
- output: `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`
- `Source unchanged: True`
- `VBA preserved: True`
- ledger installed: yes
- per-row `Επαναφορά`: **native green**

Κρατάμε το χρήσιμο visual behavior στο source slot: ο moved patient παραμένει ορατός με ένδειξη προς τη νέα θέση.

Selected/all restore παραμένει open για multi-change preview.

### 3. THERAPIST_DAILY native cleanup

Από το V82 καταγράφηκαν συγκεκριμένα cleanup items:

- η αριστερή στήλη ώρας εμφανίζει `###` σε ορισμένα rows, άρα χρειάζεται width + fixed `hh:mm`,
- το orange font είναι αποκλειστικά robotic marker,
- ο ΜΕΝΤΗΣ είναι infectious και όχι robotic, άρα δεν πρέπει να εμφανίζεται orange,
- infectious styling πρέπει να διατηρεί σωστή precedence σε moved patients,
- παλιό δεξί κουμπί `ΑΝΑΝΕΩΣΗ` από legacy VBA πρέπει να αφαιρεθεί.

### 4. Ορφανοί: operational queue με τελικό outcome

Η λίστα `Ορφανοί` είναι core daily workflow, όχι απλή ένδειξη.

Κάθε unresolved/displaced patient για τη συγκεκριμένη ημερομηνία πρέπει να καταλήγει υποχρεωτικά σε ένα από:

- `Τοποθετήθηκε` σε έγκυρο therapist/time,
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Rules:

- κανένας ορφανός δεν εξαφανίζεται σιωπηλά,
- `Αναστολή θεραπείας` είναι daily overlay/status και όχι destructive recurring edit,
- πρέπει να τροφοδοτεί τα daily statistics ως μη πραγματοποιηθείσα θεραπεία με αυτή την αιτία,
- orphan resolution συμμετέχει στο ίδιο recomputable daily engine,
- displaced existing patients προηγούνται των νέων ασθενών.

### 5. Dynamic daily recomputation

Locked order κάθε refresh/recompute:

1. recurring/base schedule για concrete date,
2. cancellations/no-shows και release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute στα net free slots,
6. resolve remaining orphans με placement ή `Αναστολή θεραπείας`,
7. μόνο μετά new-patient allocation.

Η σειρά εισαγωγής των facts δεν πρέπει να αλλάζει το τελικό valid state. Η λειτουργία είναι idempotent.

### 6. Replacement operational load semantics

Κλειδωμένο:

- authoritative load = occupied **timeslots**,
- shared slot με πολλούς patients = ένα occupied timeslot,
- ranking πρώτα με occupied timeslots, μετά session count,
- output: `load x/y | patients n`,
- example confirmed native: Μηλιδάκης `load 5/6 | patients 6`.

### 7. Mixed inpatient/outpatient presentation

Mixed effective cell:

- δεν μπλοκάρει valid movement,
- δεν παίρνει λάθος whole-cell outpatient fill,
- μπορεί προσωρινά να μείνει neutral μέχρι per-patient styling.

### 8. Outpatient recurring create/edit

Παραμένει open και επιστρέφει αμέσως μετά το current daily-state/ledger integration milestone.

## Τι έχει ολοκληρωθεί

- MASTER operational home + accepted top toolbar/header pinning στο V81.
- DAILY_INPUT authoritative flow.
- Patient-centric multi-treatment scheduling.
- Edit-existing-program preload + atomic resave.
- Conflict validation + explicit override.
- Optional manual provider/name persistence για Ρομποτικό / Εργο / Λογο / Ψυχ.
- Retired single-treatment cleanup.
- Full regression after cleanup: **699 passed, 19 warnings**.
- Cancellation slot release source support.
- Reversible ledger source gate: **10 passed in 3.26s**.
- Replacement load + mixed-slot gate: **31 passed in 2.04s**.
- V82 real preview build/install: **green**.
- V82 individual restore: **native green**.

## Τρέχον verification gate

1. Wire MASTER timetable to concrete-date `DailySessionState`.
2. Build fresh preview with active ΜΕΝΤΗΣ move.
3. Verify MASTER shows `Μηλιδάκης 10:45` while move active.
4. Verify THERAPIST_DAILY still shows source trace + destination insertion.
5. Press individual restore.
6. Verify both MASTER and THERAPIST_DAILY return to `Σκαρώνης 09:15`.
7. Fix left time-column display.
8. Fix robotic-orange / infectious styling precedence.
9. Remove legacy `ΑΝΑΝΕΩΣΗ` button.
10. Implement/verify orphan outcomes + statistics.
11. Test selected/all restore for multiple changes.
12. Replace snapshot restore with authoritative overlay removal + full daily recompute.
13. Resume outpatient recurring create/edit gate.

## Business rules που δεν πρέπει να παλινδρομήσουν

- ΦΘ provider required + solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, όχι solver constraint για inpatient flow.
- Outpatients cannot receive Ρομποτικό.
- Ανακλινόμενο / Πισίνα / ΕΦΑ χωρίς provider/name.
- Nonblank manual constraints δεν αντικαθίστανται σιωπηλά.
- Standard duration 45', Ανακλινόμενο 30'.
- Exact adjacency allowed.
- `PATIENT_PLANNER` authoritative για inpatient recurring, `OUTPATIENT_SCHEDULE` για outpatient recurring.
- Daily cancellations/absences/replacements/suspensions/redistributions είναι overlays, όχι recurring edits.
- Existing displaced patients προηγούνται των νέων ασθενών.
- Orphans πρέπει να τοποθετηθούν ή να πάρουν explicit `Αναστολή θεραπείας`.
- Daily restore αφαιρεί overlay + recompute, δεν γράφει recurring schedule από UI text.
- MASTER και THERAPIST_DAILY πρέπει να συμφωνούν στο effective daily state.
- Operational capacity είναι timeslot-based.
- Mixed inpatient/outpatient presentation ambiguity δεν ακυρώνει valid schedule movement.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
