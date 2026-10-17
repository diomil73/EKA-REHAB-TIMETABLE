# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Reversible daily change ledger: native UX prototype

Το τρέχον προσωρινό focus είναι να δούμε στην πράξη το ημερήσιο change ledger μέσα στο `THERAPIST_DAILY` πριν συνδεθεί με το authoritative recomputation engine.

Υλοποιήθηκε πρώτο native preview slice:

- `Python/rehab_excel/daily_change_ledger_vba.py`,
- `Tests/test_daily_change_ledger_vba.py`,
- integration στο υπάρχον `Python/tools/apply_replacement_choice.py`.

Το guided replacement preview πλέον μπορεί να παράγει μία πραγματική προσωρινή μετακίνηση και να εμφανίζει από κάτω:

`Αλλαγές για την dd/mm/yyyy`

με ασθενή, αρχική θέση, νέα θέση, αιτία και κατάσταση.

Prototype controls:

- `Επαναφορά` ανά row,
- checkbox selection,
- `Επαναφορά επιλεγμένων`,
- `Επαναφορά όλων`.

Prototype safety:

- metadata/snapshots αποθηκεύονται σε very-hidden `DAILY_CHANGE_LEDGER`,
- restore ενεργεί μόνο πάνω στο preview copy,
- exact source-cell snapshots χρησιμοποιούνται για native visual restoration,
- πριν από restore ελέγχεται ότι τα affected cells δεν έχουν μεταβληθεί ξανά,
- stale restore μπλοκάρεται αντί να πατάει νεότερη αλλαγή.

Focused source gate: **10 passed in 3.26s**.

Άρα το επόμενο gate είναι native smoke σε πραγματικό guided replacement preview. Μετά την αποδοχή του layout/behavior, τα ίδια controls θα συνδεθούν με removal/reversal των πραγματικών date-specific overlays και full daily recompute.

### Outpatient recurring schedule: create/edit end to end

Το patient-centric multi-treatment milestone και το retired single-treatment cleanup έχουν κλείσει πλήρως.

Confirmed green:

- edit-existing-program end to end,
- existing program preload,
- atomic resave χωρίς false self-conflict,
- optional manual names persistence/reload για Ρομποτικό / Εργο / Λογο / Ψυχ,
- MASTER fixed top toolbar + treatment headers στο V81,
- full suite πριν το cleanup: **722 passed, 19 warnings**,
- full suite μετά το cleanup: **699 passed, 19 warnings**.

Το outpatient recurring create/edit work παραμένει open και συνεχίζεται μόλις κλείσει ή προσαρμοστεί το ledger UX spike.

### Cancellation slot release

Το source-level cancellation release είναι ήδη green:

- cancelled occurrences αφαιρούνται από operational therapist/student workload,
- availability checks αγνοούν cancelled occurrences,
- replacement candidate capacity/availability λαμβάνει τα cancellations,
- dedicated regression καλύπτει workload release και exact-slot reuse.

### Κλειδωμένη λογική ημερήσιου operational recomputation

Οι ακυρώσεις εσωτερικών και εξωτερικών ασθενών και οι απουσίες θεραπευτών πρέπει να λειτουργούν ανεξάρτητα από τη χρονική σειρά με την οποία καταχωρούνται.

Η ημερήσια κατάσταση πρέπει να ξαναϋπολογίζεται από τα τελευταία γνωστά facts κάθε φορά που αλλάζει cancellation ή therapist absence:

1. recurring/base schedule για τη συγκεκριμένη ημερομηνία,
2. εφαρμογή όλων των patient cancellations / no-shows και απελευθέρωση των slots,
3. εφαρμογή όλων των therapist absences και αφαίρεση της διαθεσιμότητας αυτών των θεραπευτών,
4. εντοπισμός ασθενών που μένουν χωρίς εκτελέσιμη ανάθεση,
5. redistribution αυτών των ήδη προγραμματισμένων/επηρεασμένων ασθενών στο καθαρό διαθέσιμο capacity,
6. μόνο το capacity που απομένει μετά το redistribution προσφέρεται σε νέους ασθενείς.

Temporal invariant:

- therapist absence μπορεί να καταχωρηθεί νωρίς και να παράγει preliminary redistribution,
- μεταγενέστερη cancellation πρέπει με refresh/recompute να ελευθερώνει αμέσως νέο slot,
- correction/removal cancellation πρέπει να ξαναδεσμεύει το slot και να επανελέγχει conflicts,
- ούτε cancellation ούτε therapist absence τροποποιεί καταστροφικά το recurring base schedule,
- η λειτουργία πρέπει να είναι idempotent και να μπορεί να ξανατρέχει όσες φορές χρειάζεται μέσα στην ίδια ημέρα.

Priority invariant:

- πρώτα εξυπηρετούνται οι υπάρχοντες ασθενείς που επηρεάζονται από απουσία θεραπευτή,
- μετά χρησιμοποιείται τυχόν υπόλοιπο capacity για νέους ασθενείς.

### Locked milestone: reversible daily change ledger στο THERAPIST_DAILY

Κάτω από τα ημερήσια προγράμματα θεραπευτών θα υπάρχει date-scoped section:

`Αλλαγές για την dd/mm/yy`

Στόχος: κάθε προσωρινή operational μετακίνηση να είναι ορατή, traceable και αναστρέψιμη χωρίς destructive αλλαγή στο recurring πρόγραμμα.

Κάθε ledger row πρέπει να περιγράφει τουλάχιστον:

- ασθενή,
- αρχικό θεραπευτή και ώρα,
- νέο/effective θεραπευτή και ώρα,
- reason/source της αλλαγής,
- stable internal change/replacement id,
- active/restored status.

Required production restore semantics:

- αφαιρεί/reverses το date-specific override,
- μετά ξανατρέχει recomputation για τη συγκεκριμένη ημερομηνία,
- δεν γράφει ποτέ το recurring/base schedule από UI text,
- είναι idempotent,
- πριν από undo ελέγχει ότι δεν υπάρχει νεότερη μετακίνηση του ίδιου session,
- μετά από restore ξαναϋπολογίζονται capacity, conflicts και replacement needs,
- bulk restore πρέπει να είναι ασφαλές και να αναφέρει conflicts αντί να αφήνει μισή/ασυνεπή κατάσταση.

Next-day invariant:

- τα daily overrides ισχύουν μόνο για τη δική τους ημερομηνία,
- την επόμενη μέρα το πρόγραμμα ξεκινά ξανά από recurring/base schedule + τα overlays της νέας ημερομηνίας,
- δεν χρειάζεται χειροκίνητο undo μόνο και μόνο για να επιστρέψει ο ασθενής στη συνηθισμένη θέση του την επόμενη ημέρα.

Architecture fit:

- το υπάρχον `DailySessionState` κρατά original therapist/time και effective therapist/time ξεχωριστά,
- το `ReplacementAssignment` είναι ήδη date-specific overlay,
- άρα το ledger/revert πρέπει να επεκτείνει το υπάρχον daily-state model και όχι να δημιουργήσει δεύτερο authoritative schedule store.

### Accepted MASTER UX

- row 2: 4 horizontal MASTER action buttons,
- row 4: treatment / column headers,
- rows 1-4 μένουν σταθερές στο vertical scroll,
- το application shell ξαναβάζει το pin κάθε φορά που ενεργοποιείται το MASTER,
- no UserForm/sidebar/CommandBars.

V81 native gate: **green**.

## Τι έχει ολοκληρωθεί

- MASTER operational home και app shell.
- DAILY_INPUT authoritative flow.
- THERAPIST_DAILY recurring projection/writeback.
- Patient / therapist / student / outpatient registration slices.
- Patient-centric multi-treatment scheduling, manual/auto hybrid, atomic save.
- Conflict validation + explicit override.
- MASTER composite renderer.
- Precise ΦΘ blocker diagnostics, adjacency and overlap diagnostics.
- Optional manual provider/name fields για Ρομποτικό / Εργο / Λογο / Ψυχ.
- V80 native reset baseline verified green.
- V81 fixed top toolbar + treatment headers verified native green.
- Retired single-treatment scheduling cleanup complete.
- Post-cleanup full regression: **699 passed, 19 warnings**.
- Cancellation/released-slot propagation source implementation + dedicated regression coverage.
- Reversible daily ledger prototype source gate: **10 passed in 3.26s**.

## Τρέχον verification gate

1. Build a real guided replacement preview from V81/current workbook state.
2. Open `THERAPIST_DAILY` and inspect the ledger section under the schedules.
3. Verify the ledger row shows the correct patient and original/effective therapist-time values.
4. Test individual `Επαναφορά`.
5. Confirm exact source/destination preview cells are restored and recurring/base schedule is untouched.
6. Rebuild and test selected/all restore once multiple preview changes are supported.
7. If UX accepted, replace preview-snapshot restore with authoritative overlay removal + daily recompute.
8. Resume outpatient recurring create/edit gate.

## Business rules που δεν πρέπει να παλινδρομήσουν

- ΦΘ provider required + solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, όχι solver constraint για inpatient patient-program flow.
- Outpatients cannot receive Ρομποτικό.
- Ανακλινόμενο / Πισίνα / ΕΦΑ χωρίς provider/name στο inpatient flow.
- Nonblank manual constraints δεν αντικαθίστανται σιωπηλά.
- One final atomic save για το inpatient patient-program flow.
- Standard duration 45', Ανακλινόμενο 30'.
- Exact adjacency allowed.
- `PATIENT_PLANNER` authoritative για inpatient recurring schedule, `OUTPATIENT_SCHEDULE` authoritative για outpatient recurring schedule.
- Cancellations/absences/replacements/manual redistributions είναι date-specific operational overlays και όχι destructive edits των recurring schedules.
- Existing displaced patients προηγούνται των νέων ασθενών στη χρήση released capacity.
- Restore operations αφαιρούν overlays και κάνουν recompute, δεν ανακατασκευάζουν recurring assignments από εμφανιζόμενο κείμενο.
- Conflict override explicit όπου υποστηρίζεται.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
