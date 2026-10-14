# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Modeless MASTER sidebar UI polish

Requirement: τα 4 βασικά MASTER actions να μένουν συνεχώς ορατά κατά το vertical scroll.

Native history:

- V74: `SplitRow = 3` + `FreezePanes = True` απέτυχε.
- V75: selection-based freeze via `A4` απέτυχε.
- V76: legacy `Application.CommandBars` δεν έδωσε ορατό sidebar.
- V77: modeless sidebar δεν έφτασε στο final workbook λόγω stale shell VBA στο preview builder.
- V78: final shell refresh verified native. Το `frmMasterSidebar` εμφανίζεται και παραμένει ορατό ενώ γίνεται scroll στο MASTER.

Άρα ο μηχανισμός θεωρείται λειτουργικά σωστός. Τρέχον θέμα είναι μόνο το UI polish των κουμπιών.

Polish στο `application_shell_vba.py`:

- same 4 actions,
- `Segoe UI` 10 bold,
- centered caption μέσω `TextAlign = 2`,
- `WordWrap = False`,
- uniform button size `146 x 36`,
- equal vertical spacing.

Implementation commit: `befd4418c456be5936e6ddffddd4526009a4f725`

Regression commit: `5cc195433709d2116c97d16e5c19c455956918a5`

Focused user-run polish gate: **45 passed in 1.34s**.

### Patient-centric multi-treatment scheduling

V73 scheduling/conflict batch native green. Μετά το sidebar polish συνεχίζουμε edit-existing-program, optional names persistence/reload και full regression.

## Τι έχει ολοκληρωθεί

- MASTER operational home και app shell.
- DAILY_INPUT authoritative flow.
- THERAPIST_DAILY recurring projection/writeback.
- Patient / therapist / student / outpatient registration slices.
- Patient-centric multi-treatment scheduling, manual/auto hybrid, atomic save.
- Conflict validation + explicit override.
- MASTER composite renderer.
- Precise ΦΘ blocker diagnostics, adjacency and overlap diagnostics.
- Optional manual provider/name fields for Ρομποτικό / Εργο / Λογο / Ψυχ.
- Final preview-shell refresh so current application shell/UserForm reaches the generated preview.
- Scroll-independent modeless MASTER sidebar mechanism verified native in V78.

## Latest verified checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- freeze gate: **39 passed in 1.71s**
- stronger freeze gate: **39 passed in 2.56s**
- command-sidebar gate: **41 passed in 2.78s**
- modeless-sidebar gate: **42 passed in 1.52s**
- final shell-refresh focused gate: **44 passed in 1.41s**
- sidebar UI-polish focused gate: **45 passed in 1.34s**

## Τρέχον verification gate

1. Build V79 από clean V48.
2. Native confirm ότι το sidebar παραμένει scroll-independent και τα captions είναι σωστά κεντραρισμένα/ομοιόμορφα.
3. Confirm ότι και τα 4 actions λειτουργούν.
4. Αν η εμφάνιση γίνει αποδεκτή, κλείνουμε το sidebar milestone.
5. Μετά resume edit-existing-program, optional names persist/reload, full regression.

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
