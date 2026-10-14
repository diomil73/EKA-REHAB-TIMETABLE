# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### MASTER: σταθερά κουμπιά + ορατά treatment headers στο vertical scroll

Μετά το V80 clean reset, ο χρήστης επιβεβαίωσε ότι το edit-existing-program flow λειτουργεί σωστά και ζήτησε ένα συγκεκριμένο UX improvement:

- καθώς κατεβαίνει η λίστα στο MASTER, να παραμένουν ορατά τα ονόματα των στηλών/θεραπειών,
- και τα 4 βασικά action buttons να ακολουθούν είτε οριζόντια είτε κάθετα.

Επιλέχθηκε η πιο απλή Excel-native λύση: **κρατάμε το υπάρχον horizontal top toolbar και παγώνουμε το πάνω τμήμα μέχρι και τη γραμμή 4**.

Layout:

- row 2: 4 MASTER buttons,
- row 4: treatment / column headers,
- rows 1-4 μένουν σταθερές στο vertical scroll,
- δεν χρησιμοποιείται UserForm/sidebar/CommandBars.

Η προηγούμενη V74/V75 λογική πάγωνε λάθος boundary και δεν κάλυπτε αυτό το ακριβές requirement. Η νέα υλοποίηση χρησιμοποιεί `SplitRow = 4` και επαναφέρει το pin κάθε φορά που ενεργοποιείται το MASTER.

Implementation commit: `0c4b25eac6378e97366b6690b3443c5d46948f53`

Regression commits:

- `2860fe43bd133e2e2bac663b422ff54f5985f2c4`
- `b05426cd9cf68b39267159d84ce3d1982c145af8`

Focused source gate: **43 passed in 1.92s**.

### Patient-centric multi-treatment scheduling

Το edit-existing-program end-to-end ελέγχθηκε native από τον χρήστη χωρίς πρόβλημα.

Επόμενα patient-program checks μετά το MASTER pinning smoke:

- optional manual names persistence/reload για Ρομποτικό / Εργο / Λογο / Ψυχ,
- full regression suite,
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

## Latest verified checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- persistent-controls iteration gates: **39 / 39 / 41 / 42 / 44 / 45 passed**
- cleanup/reset focused gate: **42 passed in 3.70s**
- V80 native reset gate: **green**
- row-4 toolbar/header pinning focused source gate: **43 passed in 1.92s**

## Τρέχον verification gate

1. Build next preview from clean V48.
2. Native confirm ότι στο vertical scroll μένουν ορατά:
   - τα 4 horizontal MASTER buttons,
   - τα row-4 treatment headers.
3. Confirm no sidebar/UserForm and no unstable split/window behavior.
4. Verify optional manual names persist/reload.
5. Run full regression suite.

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
