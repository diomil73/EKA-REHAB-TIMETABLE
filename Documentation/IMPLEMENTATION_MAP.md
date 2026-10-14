# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Cleanup retired single-treatment scheduling code

Το patient-centric multi-treatment milestone είναι πλέον native green και πέρασε full regression πριν το cleanup.

Confirmed green:

- edit-existing-program end to end,
- existing program preload,
- atomic resave χωρίς false self-conflict,
- optional manual names persistence/reload για Ρομποτικό / Εργο / Λογο / Ψυχ,
- MASTER fixed top toolbar + treatment headers στο V81,
- full suite πριν το cleanup: **722 passed, 19 warnings**.

Cleanup που έγινε:

- αφαιρέθηκε `Python/rehab_excel/patient_scheduling_flow.py`,
- αφαιρέθηκε `Python/rehab_excel/patient_scheduling_form_vba.py`,
- αφαιρέθηκαν τα retired single-treatment bridge commands `plan`, `validate`, `plan-vba`, `validate-vba`, `recurring-vba`, `validate-recurring-vba`,
- κρατήθηκαν μόνο `auto-propose` / `auto-propose-vba` στο `patient_scheduling_bridge.py`, επειδή χρησιμοποιούνται από το σημερινό joint patient-program auto scheduler,
- αφαιρέθηκαν τα tests που αφορούσαν αποκλειστικά το retired single-treatment form/flow,
- το bridge regression test πλέον ελέγχει το joint auto path.

Cleanup commits:

- `93e2759726f002dd86c82168e6d3abfcf40a5d32`
- `ccc97c6f9e43cf32ca085cb42ba03f12469acb50`
- `8c0bdd326e6727c04594a6284e0148f6846bdf74`
- `8265d4000036a15323d0a0cb4ec95d6cc340e0f3`
- `29a9221055f2057f9062af7f8853274271724ef2`
- `b0d40de9f3e0eb1b76ce9cca5b1a36aa1bda3c5f`

### Accepted MASTER UX

- row 2: 4 horizontal MASTER action buttons,
- row 4: treatment / column headers,
- rows 1-4 μένουν σταθερές στο vertical scroll,
- το application shell ξαναβάζει το pin κάθε φορά που ενεργοποιείται το MASTER,
- no UserForm/sidebar/CommandBars.

Implementation commit: `0c4b25eac6378e97366b6690b3443c5d46948f53`

Regression commits:

- `2860fe43bd133e2e2bac663b422ff54f5985f2c4`
- `b05426cd9cf68b39267159d84ce3d1982c145af8`

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
- Pre-cleanup full regression: **722 passed, 19 warnings**.

## Latest verified checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- cleanup/reset focused gate: **42 passed in 3.70s**
- row-4 toolbar/header pinning focused source gate: **43 passed in 1.92s**
- V81 native pinning gate: **green**
- optional manual names persistence/reload: **native green**
- pre-cleanup full regression: **722 passed, 19 warnings**

## Τρέχον verification gate

1. Pull latest branch.
2. Run `python -m pytest` after cleanup.
3. If green, accept retired-code cleanup.
4. Update final project status / handoff.
5. Move to the next roadmap milestone.

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
