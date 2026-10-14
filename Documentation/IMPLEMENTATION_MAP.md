# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Patient-centric multi-treatment scheduling: optional manual names persistence/reload

Το MASTER UX milestone έκλεισε επιτυχώς στο V81.

Accepted MASTER behavior:

- row 2: 4 horizontal MASTER action buttons,
- row 4: treatment / column headers,
- rows 1-4 μένουν σταθερές στο vertical scroll,
- το application shell ξαναβάζει το pin κάθε φορά που ενεργοποιείται το MASTER,
- no UserForm/sidebar/CommandBars.

Implementation commit: `0c4b25eac6378e97366b6690b3443c5d46948f53`

Regression commits:

- `2860fe43bd133e2e2bac663b422ff54f5985f2c4`
- `b05426cd9cf68b39267159d84ce3d1982c145af8`

Focused source gate: **43 passed in 1.92s**.

V81 native gate: **green**. Τα 4 κουμπιά και τα treatment headers παραμένουν ορατά στο vertical scroll χωρίς unstable sidebar/window behavior.

### Patient-program current milestone

Το edit-existing-program end-to-end έχει ήδη επαληθευτεί native από τον χρήστη.

Τρέχον επόμενο checkpoint:

- optional manual names persistence/reload για Ρομποτικό / Εργο / Λογο / Ψυχ,
- typed optional names να επιβιώνουν στο save,
- να προφορτώνονται σωστά όταν γίνεται edit στο ίδιο patient program,
- μετά full regression suite,
- cleanup retired single-treatment scheduling code.

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
- Edit-existing-program native flow verified green by the user.
- V81 fixed top toolbar + treatment headers verified native green.

## Latest verified checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- persistent-controls iteration gates: **39 / 39 / 41 / 42 / 44 / 45 passed**
- cleanup/reset focused gate: **42 passed in 3.70s**
- V80 native reset gate: **green**
- row-4 toolbar/header pinning focused source gate: **43 passed in 1.92s**
- V81 native pinning gate: **green**

## Τρέχον verification gate

1. Verify optional manual names persist/reload for Ρομποτικό / Εργο / Λογο / Ψυχ.
2. Confirm names survive save and preload on edit.
3. Run full regression suite.
4. Cleanup retired single-treatment scheduling code.

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
