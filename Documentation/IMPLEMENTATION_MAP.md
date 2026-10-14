# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Modeless MASTER sidebar + final shell refresh

Requirement: τα 4 βασικά MASTER actions να μένουν συνεχώς ορατά κατά το vertical scroll.

Native history:

- V74: `SplitRow = 3` + `FreezePanes = True` απέτυχε.
- V75: selection-based freeze via `A4` απέτυχε.
- V76: legacy `Application.CommandBars` δεν έδωσε ορατό sidebar.
- V77: modeless `frmMasterSidebar` δεν εμφανίστηκε, και Alt+F8 έδειξε ότι το `ShowMasterSidebar` δεν υπήρχε καν στο τελικό preview.

Root cause V77: το patient-scheduling preview builder δεν επανεγκαθιστούσε το τρέχον `modApplicationShell` στο final output, άρα το clean V48 άφηνε stale shell VBA.

Fix:

- `build_patient_scheduling_slice_preview.py` κάνει πλέον final explicit `install_application_shell(vbproject, workbook=workbook)`,
- verifies `modApplicationShell`, `frmMasterSidebar`, και marker `Public Sub ShowMasterSidebar()`,
- μόνο μετά αποθηκεύει το preview.

Implementation commit: `9d3296a0cb8600b77a8cf1b6187494993a711cb7`

Regression: `Tests/test_patient_scheduling_preview_shell_refresh.py`

Test commit: `e5851cc3b0b1f1db21679a70bbae37994de5d704`

Focused user-run gate after the pipeline fix: **44 passed in 1.41s**.

### Patient-centric multi-treatment scheduling

V73 scheduling/conflict batch native green. Μετά το sidebar milestone συνεχίζουμε edit-existing-program, optional names persistence/reload και full regression.

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

## Τρέχον verification gate

1. Build V78 από clean V48.
2. Alt+F8 και confirm ότι υπάρχει `ShowMasterSidebar`.
3. Αν υπάρχει, verify ότι το modeless sidebar εμφανίζεται στο MASTER.
4. Confirm ότι μένει ακίνητο στο vertical scroll.
5. Confirm και τα 4 actions, hide σε άλλο sheet, restore στην επιστροφή.
6. Μετά resume edit-existing-program, optional names persist/reload, full regression.

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
