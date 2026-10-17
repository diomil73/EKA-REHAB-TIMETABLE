# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Reversible daily change ledger: native UX prototype

Το τρέχον focus είναι native έλεγχος του ημερήσιου change ledger μέσα στο `THERAPIST_DAILY` πριν συνδεθεί με το authoritative recomputation engine.

Υλοποιήθηκε πρώτο native preview slice:

- `Python/rehab_excel/daily_change_ledger_vba.py`,
- `Tests/test_daily_change_ledger_vba.py`,
- integration στο `Python/tools/apply_replacement_choice.py`.

Το guided replacement preview παράγει πραγματική προσωρινή μετακίνηση και εμφανίζει από κάτω:

`Αλλαγές για την dd/mm/yyyy`

με ασθενή, αρχική θέση, νέα θέση, αιτία και κατάσταση.

Prototype controls:

- `Επαναφορά` ανά row,
- checkbox selection,
- `Επαναφορά επιλεγμένων`,
- `Επαναφορά όλων`.

Prototype safety:

- metadata/snapshots σε very-hidden `DAILY_CHANGE_LEDGER`,
- restore μόνο πάνω στο preview copy,
- exact source-cell snapshots για native visual restoration,
- stale restore μπλοκάρεται πριν πατήσει νεότερη αλλαγή.

Focused ledger source gate: **10 passed in 3.26s**.
Replacement load + mixed-slot regression gate: **31 passed in 2.04s**.

### V82 real guided replacement build

Το πραγματικό stress test χτίστηκε επιτυχώς:

- date: `2026-10-19`
- patient: `ΜΕΝΤΗΣ`
- original: `Σκαρώνης 09:15`
- replacement: `Μηλιδάκης 10:45`
- Μηλιδάκης candidate: `load 5/6 | patients 6`
- output: `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`
- `Source unchanged: True`
- `VBA preserved: True`
- `Daily change ledger installed: yes`

Άρα build/install gate: **green**. Το επόμενο gate είναι visual/interactive native smoke μέσα στο `THERAPIST_DAILY`.

### Replacement operational load semantics

Κλειδώθηκε ότι το operational load είναι τα κατειλημμένα **timeslots** και όχι ο ακατέργαστος αριθμός sessions/patients.

- `active_timeslots` = authoritative load/capacity.
- shared slot με πολλούς ασθενείς μετρά ως ένα occupied timeslot.
- candidate ranking ταξινομεί πρώτα με occupied timeslots και μετά με session count.
- candidate output διαχωρίζει `load x/y` από `patients n`.
- παράδειγμα: 6 sessions σε 5 διαφορετικά slots => `load 5/6 | patients 6`.

### Mixed inpatient/outpatient slot presentation

Mixed effective cell δεν πρέπει να ακυρώνει έγκυρη operational μετακίνηση.

- δεν εφαρμόζεται ενιαίο outpatient whole-cell fill σε mixed cell,
- το replacement preview συνεχίζει κανονικά,
- το mixed cell μένει ουδέτερο μέχρι να έχουμε per-patient text styling.

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

Η ημερήσια κατάσταση ξαναϋπολογίζεται από τα τελευταία γνωστά facts:

1. recurring/base schedule για τη συγκεκριμένη ημερομηνία,
2. patient cancellations/no-shows και release slots,
3. therapist absences,
4. displaced/unscheduled existing patients,
5. redistribution στο καθαρό διαθέσιμο capacity,
6. μόνο μετά new-patient allocation.

Temporal invariant:

- absence μπορεί να προηγηθεί cancellation,
- μεταγενέστερη cancellation πρέπει να βελτιώνει το recomputed plan,
- correction/removal cancellation ξαναδεσμεύει slot και απαιτεί revalidation,
- ούτε cancellation ούτε absence γράφει destructively στο recurring schedule,
- recomputation idempotent.

Priority invariant:

- πρώτα displaced existing patients,
- μετά new patients.

### Locked milestone: reversible daily change ledger στο THERAPIST_DAILY

Κάτω από τα ημερήσια προγράμματα θεραπευτών υπάρχει date-scoped section:

`Αλλαγές για την dd/mm/yy`

Κάθε ledger row πρέπει να περιγράφει τουλάχιστον:

- ασθενή,
- αρχικό θεραπευτή και ώρα,
- νέο/effective θεραπευτή και ώρα,
- reason/source,
- stable internal change/replacement id,
- active/restored status.

Required production restore semantics:

- αφαιρεί/reverses το date-specific override,
- μετά full recompute για τη συγκεκριμένη ημερομηνία,
- δεν γράφει ποτέ recurring/base schedule από UI text,
- idempotent,
- stale undo μπλοκάρεται,
- capacity/conflicts/replacements ξαναϋπολογίζονται,
- bulk restore ασφαλές και atomic/validated.

Next-day invariant:

- daily overrides ισχύουν μόνο για τη δική τους ημερομηνία,
- η επόμενη μέρα ξεκινά από recurring/base + overlays της νέας ημέρας,
- δεν χρειάζεται manual undo για επιστροφή στο recurring πρόγραμμα.

Architecture fit:

- `DailySessionState` κρατά original και effective therapist/time ξεχωριστά,
- `ReplacementAssignment` είναι date-specific overlay,
- το ledger/revert επεκτείνει το daily-state model και δεν δημιουργεί δεύτερο authoritative schedule store.

### Accepted MASTER UX

- row 2: 4 horizontal MASTER action buttons,
- row 4: treatment / column headers,
- rows 1-4 μένουν σταθερές στο vertical scroll,
- application shell reapplies pin,
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
- Replacement operational-load semantics + mixed-slot presentation fix: **31 passed in 2.04s**.
- V82 real guided replacement preview build/install: **green**.

## Τρέχον verification gate

1. Open `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`.
2. Go to `THERAPIST_DAILY`.
3. Confirm ΜΕΝΤΗΣ appears moved from Σκαρώνης 09:15 to Μηλιδάκης 10:45.
4. Scroll below schedules and verify `Αλλαγές για την 19/10/2026`.
5. Verify row details are correct.
6. Test individual `Επαναφορά`.
7. Confirm source/destination preview cells restore exactly.
8. Confirm ledger row stays visible as restored history.
9. Confirm recurring/base schedule remains untouched.
10. Later test selected/all restore with multiple changes.
11. After UX acceptance, replace snapshot restore with authoritative overlay removal + daily recompute.
12. Resume outpatient recurring create/edit gate.

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
- Replacement operational load/capacity είναι timeslot-based και όχι raw session-count based.
- Mixed inpatient/outpatient cell δεν μπλοκάρει valid operational movement μόνο λόγω whole-cell presentation ambiguity.
- Conflict override explicit όπου υποστηρίζεται.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
