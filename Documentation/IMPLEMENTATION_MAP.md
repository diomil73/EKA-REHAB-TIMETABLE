# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Patient-centric multi-treatment scheduling: edit-existing-program

Το persistent-controls/sidebar detour έκλεισε. Το V80 clean reset preview είναι native green και το MASTER επέστρεψε σε σταθερό baseline με το υπάρχον top toolbar.

Τρέχον milestone:

- Βασικό μενού → **Επεξεργασία ασθενή**,
- existing patient registry edit,
- άνοιγμα της ίδιας patient-program φόρμας,
- preload του υπάρχοντος προγράμματος,
- αλλαγή treatment/day/time/provider,
- atomic resave,
- verification σε `PATIENT_PLANNER` και MASTER composite rendering,
- χωρίς false self-conflicts από τις προηγούμενες persisted assignments του ίδιου patient.

Μετά ακολουθούν:

- optional manual names persistence/reload για Ρομποτικό / Εργο / Λογο / Ψυχ,
- full regression suite,
- cleanup retired single-treatment scheduling code.

### MASTER persistent-controls status

Native history:

- V74/V75 freeze approaches απέτυχαν.
- V76 legacy CommandBars απέτυχε.
- V77 αποκάλυψε stale application-shell VBA στο final preview builder.
- V78 επιβεβαίωσε το final shell-refresh fix.
- V79 απέρριψε το modeless UserForm ως μη αξιόπιστο operational UI.
- V80 clean reset είναι native green: no sidebar/UserForm, no odd freeze/split residue, normal top toolbar/navigation/Save & Exit λειτουργούν σωστά.

Απόφαση: workbook-only sidebar attempts retired. Πραγματικό docked sidebar απαιτεί ξεχωριστό Excel add-in/custom task pane milestone.

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
- Cleanup των αποτυχημένων sidebar/freeze runtime mechanisms.
- V80 native reset baseline verified green.

## Latest verified checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- persistent-controls iteration gates: **39 / 39 / 41 / 42 / 44 / 45 passed**
- cleanup/reset focused gate: **42 passed in 3.70s**
- V80 native reset gate: **green**

## Τρέχον verification gate

1. Native edit-existing-program end to end.
2. Confirm existing program preload.
3. Modify one or more assignments and save atomically.
4. Verify `PATIENT_PLANNER` + MASTER after the edit.
5. Confirm no false self-conflict from replaced persisted assignments.
6. Verify optional manual names persist/reload.
7. Run full regression suite.

## Business rules που δεν πρέπει να παλινδρομήσουν

- ΦΘ provider required + solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, όχι solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ χωρίς provider/name.
- Nonblank manual constraints δεν αντικαθίστανται σιωπηρά.
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
