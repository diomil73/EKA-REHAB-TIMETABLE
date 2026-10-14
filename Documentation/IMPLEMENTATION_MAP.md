# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ο ζωντανός χάρτης υλοποίησης του project. Για κάθε ουσιαστικό milestone καταγράφει: **τι φτιάχνουμε τώρα**, **τι έχει ολοκληρωθεί**, **ποιο είναι το verification gate** και **ποια είναι τα επόμενα βήματα**.

Για το ακριβές σημείο συνέχισης μετά από pause/context loss, πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

---

## Τι φτιάχνουμε τώρα

### Modeless MASTER sidebar

Ο χρήστης ζήτησε τα βασικά κουμπιά του MASTER να παραμένουν πάντα ορατά κατά το scroll ή να μεταφερθούν σε αριστερό sidebar.

Έχουν ήδη απορριφθεί native τρεις μηχανισμοί:

- V74: `SplitRow = 3` + `FreezePanes = True`,
- V75: selection-based freeze μέσω `A4`,
- V76: legacy `Application.CommandBars` sidebar.

Οι δύο freeze εκδοχές δεν κράτησαν τα worksheet shape-buttons ορατά στο vertical scroll. Το command-bar sidebar δεν παρήγαγε ορατή native αλλαγή στο σύγχρονο Excel/app-mode setup.

Η νέα λύση είναι **modeless VBA UserForm** `frmMasterSidebar`, ανεξάρτητο από cells, scroll, freeze panes και legacy command bars.

Η νέα υλοποίηση στο `application_shell_vba.py`:

- εγκαθιστά το `frmMasterSidebar` στο VBA project,
- εμφανίζεται modeless μόνο στο `MASTER_SCHEDULE`,
- τοποθετείται κοντά στο αριστερό άκρο του Excel window,
- περιέχει τις τέσσερις ενέργειες **Διαχείριση Μητρώων**, **Απουσίες / Ακυρώσεις**, **Ημερήσιο πρόγραμμα**, **Save & Exit**,
- κρύβει τα legacy `ekaToolbar_` shapes όσο το sidebar είναι ορατό,
- κρύβεται όταν ο χρήστης πάει σε άλλο sheet,
- unload στο application-shell exit.

Implementation commit: `decf535fa90a4ffd4d645efc0529762c72238744`

Regression update: `dfb5a38ead7ea375811459abe45ffc458fd2e1fc`

Focused user-run gate: **42 passed in 1.52s**.

Επόμενο gate: **V77 από clean V48** και native verification του modeless sidebar.

### Patient-centric multi-treatment scheduling

Το V73 scheduling/conflict batch είναι native green. Μετά το sidebar UX, επιστρέφουμε στο επόμενο patient-program milestone: edit-existing-program flow, optional manual names και full regression.

Κύριοι κανόνες που παραμένουν:

- ΦΘ: provider υποχρεωτικός και solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ: optional free-text όνομα, όχι solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ: χωρίς provider/name.
- Manual nonblank constraints δεν αντικαθίστανται σιωπηρά.
- Standard duration 45', Ανακλινόμενο 30'.
- Exact adjacency επιτρέπεται.
- `PATIENT_PLANNER` authoritative, `MASTER_SCHEDULE` derived view.

---

## Τι έχει ολοκληρωθεί

### Application / workbook foundation

- `MASTER_SCHEDULE` operational home.
- App-mode shell και βασική πλοήγηση.
- `DAILY_INPUT` authoritative flow.
- `THERAPIST_DAILY` recurring projection/writeback.
- Patient / therapist / student / outpatient registration vertical slices.
- Outpatient scheduling και unified base schedule compatibility.
- MASTER projection, room sorting, infectious/outpatient/robotic visual rules.

### Patient scheduling architecture

- Patient-centric draft πολλών θεραπειών.
- Dynamic treatments από workbook SETTINGS.
- Manual + partial-auto + fully-auto flow.
- Atomic final save με snapshots, verification και rollback.
- CURRENT preload από `PATIENT_PLANNER`.
- Ίδια patient-program φόρμα μετά από **Επεξεργασία ασθενή**.
- Optional manual names για Ρομποτικό / Εργο / Λογο / Ψυχ.
- Duration-aware patient conflicts και ΦΘ/Ανακλινόμενο adjacency scoring.
- MASTER composite renderer.

### Verified native checkpoints

- V69: first save + MASTER composite rendering.
- V70: ΦΘ self-conflict false alarm fixed.
- V71: installer/conflict flow native operational.
- V73: precise ΦΘ blocker diagnostics, exact adjacency, detailed overlap diagnostics, conflict-dialog layout και normal Ρομποτικό rendering verified native.
- V74: toolbar actions work, freeze visibility failed.
- V75: stronger freeze attempt still failed visibility; freeze direction retired.
- V76: command-bar sidebar produced no visible native change; legacy CommandBars direction retired.

### Source/integration checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- conflict layout + MASTER refresh gate: **57 passed in 3.96s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- first persistent MASTER toolbar gate: **39 passed in 1.71s**
- stronger selection-based freeze gate: **39 passed in 2.56s**
- command-sidebar gate: **41 passed in 2.78s**
- modeless UserForm sidebar gate: **42 passed in 1.52s**

### Conflict-validation / Save override

- Auto και Save χρησιμοποιούν κοινό conflict validation.
- Standard 45', Ανακλινόμενο 30'.
- Shared weekdays only.
- Exact adjacency allowed.
- Save conflict dialog έχει **Συνέχεια / Επιστροφή**.
- **Επιστροφή** σημαίνει zero writes.
- **Συνέχεια** σημαίνει explicit override.

### Edit-draft replacement + diagnostics

- persisted assignments του ίδιου patient για treatments που αντικαθίστανται στο draft δεν μπλοκάρουν το νέο draft,
- οι νέες assignments ελέγχονται joint,
- manual ΦΘ occupancy diagnostic μπορεί να κατονομάσει weekday/time/provider/blocking patient,
- true overlaps εμφανίζουν treatments + ranges.

---

## Τρέχον verification gate

Για το modeless MASTER sidebar:

1. build `MASTER_TOOLBAR_SMOKE_V77.xlsm` από clean V48,
2. native επιβεβαίωση ότι το modeless sidebar εμφανίζεται στο MASTER και δεν μετακινείται με vertical scroll,
3. έλεγχος και των τεσσάρων actions,
4. μετάβαση σε άλλο sheet και επιβεβαίωση hide,
5. επιστροφή στο MASTER και επιβεβαίωση restore.

Το focused source gate είναι πράσινο: **42 passed in 1.52s**.

Μετά συνεχίζει το patient-program gate:

- edit-existing-program end to end,
- optional manual names persist/reload,
- full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Build V77 από clean V48.
2. Native smoke του modeless MASTER sidebar.
3. Αν είναι πράσινο, επιστροφή στο edit-existing-program milestone.
4. Optional manual names persist/reload.
5. Full regression suite.

### Μετά το patient-program milestone

- cleanup retired single-treatment scheduling code,
- τελικός UX pass,
- ενημέρωση `UI_APPLICATION_ROADMAP.md` και `PROJECT_STATUS.md`,
- μετάβαση στο επόμενο ανοιχτό roadmap milestone.

---

## Recovery rule

Αν χαθεί context, διάβασε με αυτή τη σειρά:

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Σε Windows χρησιμοποιούμε πάντα `python -m pytest`, ποτέ bare `pytest`.
