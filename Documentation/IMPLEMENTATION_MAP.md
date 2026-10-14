# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ο ζωντανός χάρτης υλοποίησης του project. Για κάθε ουσιαστικό milestone καταγράφει: **τι φτιάχνουμε τώρα**, **τι έχει ολοκληρωθεί**, **ποιο είναι το verification gate** και **ποια είναι τα επόμενα βήματα**.

Για το ακριβές σημείο συνέχισης μετά από pause/context loss, πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

---

## Τι φτιάχνουμε τώρα

### Scroll-independent MASTER sidebar

Ο χρήστης ζήτησε τα βασικά κουμπιά του MASTER να παραμένουν πάντα ορατά κατά το scroll ή να μεταφερθούν σε αριστερό sidebar.

Δοκιμάστηκαν δύο freeze-based εκδοχές:

- V74: `SplitRow = 3` + `FreezePanes = True`,
- V75: selection-based freeze μέσω `A4` + `FreezePanes = True`.

Και οι δύο πέρασαν source tests, αλλά απέτυχαν native στο βασικό requirement: τα worksheet shape-buttons εξακολουθούν να φεύγουν με το vertical scroll. Τα actions/navigation παρέμειναν λειτουργικά.

Η freeze κατεύθυνση εγκαταλείφθηκε. Η νέα λύση είναι πραγματικό αριστερό Excel command sidebar, ανεξάρτητο από τα cells και το scroll.

Η νέα υλοποίηση στο `application_shell_vba.py`:

- δημιουργεί προσωρινό command bar `EKA_MASTER_SIDEBAR`, docked αριστερά,
- έχει τις τέσσερις ενέργειες **Διαχείριση Μητρώων**, **Απουσίες / Ακυρώσεις**, **Ημερήσιο πρόγραμμα**, **Save & Exit**,
- εμφανίζεται μόνο στο `MASTER_SCHEDULE`,
- κρύβεται στα άλλα operational sheets,
- κρύβει τα legacy `ekaToolbar_` shapes όταν είναι ενεργό ώστε να μην υπάρχουν διπλά controls,
- διαγράφεται στο exit του application shell.

Implementation commit: `ad7bfbd7045a445eddb747d5b5f91d482c6046b0`

Regression commits:

- `37f38b41df2f9a3d18cfba04046b7c45bbefcf4d`
- `5fcbc3159a7c603a556be788a5596fd69fd34e34`

Focused user-run gate: **41 passed in 2.78s**.

Επόμενο gate: **V76 από clean V48** και native verification του αριστερού sidebar.

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

### Source/integration checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- conflict layout + MASTER refresh gate: **57 passed in 3.96s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- first persistent MASTER toolbar gate: **39 passed in 1.71s**
- stronger selection-based freeze gate: **39 passed in 2.56s**
- command-sidebar gate: **41 passed in 2.78s**

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

Για το MASTER sidebar:

1. build `MASTER_TOOLBAR_SMOKE_V76.xlsm` από clean V48,
2. native επιβεβαίωση ότι το αριστερό sidebar μένει ακίνητο ενώ γίνεται vertical scroll,
3. έλεγχος και των τεσσάρων actions,
4. μετάβαση σε άλλο sheet και επιβεβαίωση hide,
5. επιστροφή στο MASTER και επιβεβαίωση restore.

Το focused source gate είναι πράσινο: **41 passed in 2.78s**.

Μετά συνεχίζει το patient-program gate:

- edit-existing-program end to end,
- optional manual names persist/reload,
- full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Build V76 από clean V48.
2. Native smoke του αριστερού MASTER sidebar.
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
