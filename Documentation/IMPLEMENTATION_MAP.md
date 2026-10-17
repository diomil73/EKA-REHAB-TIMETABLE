# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Ενοποίηση daily effective state σε MASTER + THERAPIST_DAILY

Locked target:

`DailySessionState -> THERAPIST_DAILY + MASTER timetable + ledger + statistics`

Το V83 πέρασε native το βασικό integration stress test:

- ΜΕΝΤΗΣ `Σκαρώνης 09:15 -> Μηλιδάκης 10:45`,
- `THERAPIST_DAILY` δείχνει σωστά active move + source trace,
- `MASTER` δείχνει επίσης το effective daily assignment,
- individual `Επαναφορά` γυρίζει και τα δύο views στο `Σκαρώνης 09:15`.

Status: **MASTER + THERAPIST_DAILY + ledger native green για το V83 stress case**. Statistics integration παραμένει open.

Focused source gate πριν το V83: **24 passed, 5 warnings in 4.39s**. Τα warnings είναι Python 3.14 `ZipFile.__del__` teardown warnings και δεν αποτελούν functional blocker.

### 2. THERAPIST_DAILY time-axis cleanup

Το V83 επιβεβαίωσε ότι οι ώρες αριστερά εξακολουθούν να εμφανίζονται ως `####`.

Root cause: `RefreshDailyChangeLedgerView` ξαναέβαζε τη στήλη A σε width 4.

Fix implemented:

- column A width 8,
- `A2:A8,A12:A18` fixed `hh:mm`,
- centered time display,
- regression απαγορεύει επιστροφή σε width 4.

Status: **implemented, focused/native verification pending**.

### 3. V82/V83 reversible daily change ledger

- per-row `Επαναφορά`: native green,
- source trace διατηρείται και είναι αποδεκτό UX,
- selected/all restore παραμένει open για multi-change preview,
- production restore αργότερα θα γίνει authoritative overlay removal + recompute αντί snapshot restore.

### 4. THERAPIST_DAILY styling / legacy cleanup

Open items:

- orange font αποκλειστικά robotic marker,
- ο ΜΕΝΤΗΣ είναι infectious και όχι robotic, άρα δεν πρέπει να εμφανίζεται orange,
- infectious styling precedence σε moved patients,
- αφαίρεση legacy δεξιού κουμπιού `ΑΝΑΝΕΩΣΗ`.

### 5. Ορφανοί: operational queue με τελικό outcome

Κάθε unresolved/displaced patient για τη συγκεκριμένη ημερομηνία πρέπει να καταλήγει σε:

- `Τοποθετήθηκε` σε έγκυρο therapist/time, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Rules:

- κανένας ορφανός δεν εξαφανίζεται σιωπηλά,
- `Αναστολή θεραπείας` είναι daily overlay/status, όχι recurring edit,
- τροφοδοτεί daily statistics ως μη πραγματοποιηθείσα θεραπεία με αυτή την αιτία,
- displaced existing patients προηγούνται των νέων ασθενών.

### 6. Dynamic daily recomputation

Locked order:

1. recurring/base schedule για concrete date,
2. cancellations/no-shows και release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute στα net free slots,
6. resolve remaining orphans με placement ή `Αναστολή θεραπείας`,
7. μόνο μετά new-patient allocation.

Η σειρά εισαγωγής facts δεν αλλάζει το τελικό valid state. Recomputation idempotent.

### 7. Replacement operational load semantics

- authoritative load = occupied **timeslots**,
- shared slot με πολλούς patients = ένα occupied timeslot,
- ranking πρώτα occupied timeslots, μετά session count,
- output `load x/y | patients n`,
- confirmed: Μηλιδάκης `load 5/6 | patients 6`.

### 8. Mixed inpatient/outpatient presentation

Mixed effective cell:

- δεν μπλοκάρει valid movement,
- δεν παίρνει λάθος whole-cell outpatient fill,
- μπορεί προσωρινά να μείνει neutral μέχρι per-patient styling.

### 9. Outpatient recurring create/edit

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
- MASTER daily overlay + restore source gate: **24 passed, 5 warnings in 4.39s**.
- V82 real preview build/install: green.
- V82 individual restore: native green.
- V83 MASTER active-overlay sync + restore: **native green**.

## Τρέχον verification gate

1. Run `Tests/test_daily_change_ledger_vba.py` after the time-axis fix.
2. Build fresh preview and verify no `####` remain in the left time column.
3. Fix robotic-orange / infectious styling precedence.
4. Remove legacy `ΑΝΑΝΕΩΣΗ` button.
5. Implement/verify orphan outcomes + statistics.
6. Test selected/all restore for multiple changes.
7. Replace snapshot restore with authoritative overlay removal + full daily recompute.
8. Resume outpatient recurring create/edit gate.
9. Clean up Python 3.14 ZipFile teardown warnings later unless they become functional failures.

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
