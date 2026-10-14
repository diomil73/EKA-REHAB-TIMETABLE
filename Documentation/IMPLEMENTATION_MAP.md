# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### Stable MASTER baseline after sidebar experiments

Το requirement ήταν τα 4 βασικά MASTER actions να μένουν συνεχώς ορατά κατά το vertical scroll.

Native ιστορικό:

- V74: `SplitRow = 3` + `FreezePanes = True` απέτυχε.
- V75: selection-based freeze μέσω `A4` απέτυχε.
- V76: legacy `Application.CommandBars` δεν έδωσε αξιόπιστο visible sidebar.
- V77: modeless sidebar δεν είχε περάσει στο final workbook λόγω stale shell VBA.
- V78: επιβεβαίωσε ότι το final shell refresh fix δουλεύει και το UserForm μπορεί να μείνει visible στο scroll.
- V79: το modeless UserForm απορρίφθηκε native λόγω ασυνεπούς εμφάνισης, Alt+F8/Alt+F4 interaction και unstable window behavior.

Απόφαση: σταματούν τα workbook-only sidebar experiments. Επιστρέφουμε σε σταθερό application shell και στο υπάρχον top MASTER toolbar. Πραγματικό docked sidebar θα απαιτούσε ξεχωριστό Excel add-in/custom task pane milestone.

Cleanup:

- αφαιρέθηκαν `ShowMasterSidebar` / `HideMasterSidebar`,
- αφαιρέθηκε το modeless UserForm runtime,
- δεν υπάρχει runtime freeze/split pinning στο application shell,
- το installer αφαιρεί ρητά το retired `frmMasterSidebar`,
- κρατήθηκε το final preview-shell refresh ώστε το generated preview να παίρνει πάντα το current `modApplicationShell`.

Cleanup commits:

- `a0158a479904c26cff8fe1f4265ad777dd8426e6`
- `c7760808f2f5ad5b5f6798a7e34521ab0e50f3f5`
- `87e6388cd73c1f62f67761d96130d1379c0904a4`
- `effadb2973326046b0c395c72f71916647549c4e`
- `a2fb0c266975c7322fe93f3733043a2a4807b2d2`

### Patient-centric multi-treatment scheduling

V73 scheduling/conflict batch είναι native green. Μόλις περάσει το cleanup gate, επιστρέφουμε στο επόμενο patient-program milestone:

- edit-existing-program end to end,
- optional manual names persistence/reload,
- full regression.

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

## Latest verified checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- persistent-controls iteration gates: **39 / 39 / 41 / 42 / 44 / 45 passed**
- latest pre-cleanup focused gate: **45 passed in 1.34s**

## Τρέχον verification gate

1. Run focused application-shell / toolbar / preview-shell-refresh tests after cleanup.
2. Αν είναι green, build ένα clean reset preview από V48.
3. Native confirm:
   - no sidebar/UserForm,
   - no odd freeze/split residue,
   - top MASTER toolbar works normally,
   - navigation and Save & Exit remain operational.
4. Resume edit-existing-program end to end.
5. Verify optional manual names persist/reload.
6. Run full regression suite.

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
