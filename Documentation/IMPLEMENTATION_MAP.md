# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

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

Η μείωση του test count είναι αναμενόμενη, επειδή αφαιρέθηκαν tests που κάλυπταν αποκλειστικά retired single-treatment code.

### Cancellation slot release

Το προηγούμενο roadmap candidate είναι ήδη source-green και δεν χρειάζεται νέα υλοποίηση:

- cancelled occurrences αφαιρούνται από operational therapist/student workload,
- availability checks αγνοούν cancelled occurrences,
- replacement candidate capacity/availability λαμβάνει τα cancellations,
- το `Tests/test_cancellation_releases_replacement_capacity.py` καλύπτει ακριβώς workload release και exact-slot reuse.

Native cancellation smoke παραμένει μελλοντικό verification item, όχι source implementation gap.

### Τρέχον outpatient milestone

Στόχος: native create/edit flow για recurring outpatient schedule στο `OUTPATIENT_SCHEDULE`.

Υπάρχει ήδη backend upsert με `OutpatientScheduleRequest.target_base_entry_id`, conflict validation, suggestions και source-hash/readback safety. Τώρα ελέγχουμε αν το Excel form διαθέτει user-facing edit/preload. Αν λείπει, θα προστεθεί πάνω στο υπάρχον backend αντί να δημιουργηθεί δεύτερος μηχανισμός.

Acceptance:

- create recurring outpatient entry,
- edit το ίδιο entry με preload,
- update της ίδιας authoritative row,
- no outpatient writes to `PATIENT_PLANNER`,
- robotic treatment unavailable,
- conflict/suggestion safeguards preserved,
- source workbook unchanged,
- VBA preserved,
- native Windows/Excel smoke green.

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
- Final preview-shell refresh ώστε το current application shell να φτάνει πάντα στο generated preview.
- Cleanup των αποτυχημένων sidebar mechanisms.
- V80 native reset baseline verified green.
- Edit-existing-program native flow verified green.
- V81 fixed top toolbar + treatment headers verified native green.
- Optional manual names save/reload verified native green.
- Retired single-treatment scheduling cleanup complete.
- Post-cleanup full regression: **699 passed, 19 warnings**.
- Cancellation/released-slot propagation source implementation + dedicated regression coverage.

## Τρέχον verification gate

1. Inspect outpatient form/bridge create-vs-edit support.
2. Add missing edit/preload UI only if needed.
3. Run focused outpatient schedule tests.
4. Build native smoke preview.
5. Verify create + edit against `OUTPATIENT_SCHEDULE` in Excel.

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
- Conflict override explicit όπου υποστηρίζεται.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
