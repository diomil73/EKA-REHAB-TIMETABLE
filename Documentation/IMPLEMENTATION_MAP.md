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

### Τρέχον outpatient milestone

Στόχος: native create/edit flow για recurring outpatient schedule στο `OUTPATIENT_SCHEDULE`.

Υπάρχει ήδη backend upsert με `OutpatientScheduleRequest.target_base_entry_id`, conflict validation, suggestions και source-hash/readback safety. Προστίθεται user-facing edit/preload πάνω στο ίδιο form/backend.

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
- V80 native reset baseline verified green.
- V81 fixed top toolbar + treatment headers verified native green.
- Retired single-treatment scheduling cleanup complete.
- Post-cleanup full regression: **699 passed, 19 warnings**.
- Cancellation/released-slot propagation source implementation + dedicated regression coverage.

## Επόμενα gates

1. Complete outpatient create/edit source wiring and focused tests.
2. Native outpatient create/edit smoke against `OUTPATIENT_SCHEDULE`.
3. Add/verify the dynamic daily recomputation contract for cancellations + therapist absences entered in either order.
4. Native daily operational smoke with scenarios:
   - absence first, cancellation later,
   - cancellation first, absence later,
   - cancellation correction/removal,
   - inpatient and outpatient cancellations,
   - redistribution before new-patient allocation.

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
- Cancellations/absences είναι date-specific operational overlays και όχι destructive edits των recurring schedules.
- Existing displaced patients προηγούνται των νέων ασθενών στη χρήση released capacity.
- Conflict override explicit όπου υποστηρίζεται.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
