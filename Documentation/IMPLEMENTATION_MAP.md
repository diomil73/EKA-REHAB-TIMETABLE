# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Επόμενο roadmap milestone

Το patient-centric multi-treatment scheduling milestone έκλεισε πλήρως, μαζί με το cleanup του retired single-treatment κώδικα.

Confirmed green:

- edit-existing-program end to end,
- existing program preload,
- atomic resave χωρίς false self-conflict,
- optional manual names persistence/reload για Ρομποτικό / Εργο / Λογο / Ψυχ,
- MASTER fixed top toolbar + treatment headers στο V81,
- full suite πριν το cleanup: **722 passed, 19 warnings**,
- full suite μετά το cleanup: **699 passed, 19 warnings**.

Η μείωση του test count είναι αναμενόμενη, επειδή αφαιρέθηκαν tests που κάλυπταν αποκλειστικά retired single-treatment code.

### Accepted MASTER UX

- row 2: 4 horizontal MASTER action buttons,
- row 4: treatment / column headers,
- rows 1-4 μένουν σταθερές στο vertical scroll,
- το application shell ξαναβάζει το pin κάθε φορά που ενεργοποιείται το MASTER,
- no UserForm/sidebar/CommandBars.

V81 native gate: **green**.

### Retired scheduling cleanup

Ολοκληρώθηκε και επαληθεύτηκε.

Αφαιρέθηκαν:

- `Python/rehab_excel/patient_scheduling_flow.py`,
- `Python/rehab_excel/patient_scheduling_form_vba.py`,
- retired bridge commands `plan`, `validate`, `plan-vba`, `validate-vba`, `recurring-vba`, `validate-recurring-vba`,
- tests αποκλειστικά για το retired single-treatment flow/form.

Διατηρήθηκαν:

- `auto-propose` / `auto-propose-vba` για το joint patient-program auto scheduler,
- το current `PatientProgramRuntime` / `frmPatientScheduling`,
- cleanup/verification ώστε παλιά workbooks να μην κρατούν `PatientSchedulingRuntime`.

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

## Τρέχον verification gate

Το προηγούμενο milestone είναι κλειστό. Το επόμενο gate θα οριστεί από το επόμενο roadmap item, με βάση τα υπόλοιπα outpatient/operational gaps στο `Documentation/PROJECT_STATUS.md`.

## Business rules που δεν πρέπει να παλινδρομήσουν

- ΦΘ provider required + solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, όχι solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ χωρίς provider/name.
- Nonblank manual constraints δεν αντικαθίστανται σιωπηλά.
- One final atomic save.
- Standard duration 45', Ανακλινόμενο 30'.
- Exact adjacency allowed.
- `PATIENT_PLANNER` authoritative, MASTER derived.
- Conflict override explicit, return-for-correction zero writes.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
