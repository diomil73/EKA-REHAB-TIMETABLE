# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

## Σκοπός

Ο ζωντανός χάρτης υλοποίησης του project. Για κάθε ουσιαστικό milestone καταγράφει: **τι φτιάχνουμε τώρα**, **τι έχει ολοκληρωθεί**, **ποιο είναι το verification gate** και **ποια είναι τα επόμενα βήματα**.

Για το ακριβές σημείο συνέχισης μετά από pause/context loss, πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

---

## Τι φτιάχνουμε τώρα

### Persistent MASTER toolbar

Ο χρήστης ζήτησε τα βασικά κουμπιά του MASTER να παραμένουν πάντα ορατά κατά το scroll ή να μεταφερθούν σε αριστερό sidebar.

Επιλέχθηκε η ασφαλέστερη εκδοχή: **παραμένουν στην υπάρχουσα επάνω μπάρα και παγώνουν οι γραμμές 1–3**. Έτσι η γραμμή 2 με τα κουμπιά μένει συνεχώς ορατή χωρίς να αλλάξει ούτε μία στήλη του MASTER.

Η υλοποίηση στο `application_shell_vba.py` προσθέτει `PinMasterToolbar()`:

- `FreezePanes = False`,
- `SplitColumn = 0`,
- `SplitRow = 3`,
- `FreezePanes = True`,
- επανεφαρμογή στο app entry και σε κάθε επιστροφή/activation του MASTER.

Source commit: `346247ea1dde7f9a3529bf6c25d786b08409699b`

Regression coverage:

- `Tests/test_master_toolbar_pinned_view.py`
- application-shell activation regression updated για το multi-line MASTER path.

Test commits:

- `0b5eb5a18a2ba0307bf57d61099d38b9d6a612fb`
- `2226f27d7738d43974c213e952c076993afd4a0c`

Focused user-run gate: **39 passed in 1.71s**.

Επόμενο gate: build + native verification του V74.

### Patient-centric multi-treatment scheduling

Το V73 scheduling/conflict batch είναι native green. Μετά το toolbar polish, επιστρέφουμε στο επόμενο patient-program milestone: edit-existing-program flow, optional manual names και full regression.

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
- Dynamic treatments από SETTINGS.
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

### Source/integration checkpoints

- self-conflict gate: **32 passed in 1.40s**
- V71 conflict/installer gate: **62 passed in 3.34s**
- conflict layout + MASTER refresh gate: **57 passed in 3.96s**
- edit-draft diagnostics gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**
- persistent MASTER toolbar gate: **39 passed in 1.71s**

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

Για το MASTER toolbar:

1. build `MASTER_TOOLBAR_SMOKE_V74.xlsm` από clean V48,
2. native vertical scroll με τα τέσσερα κουμπιά συνεχώς ορατά,
3. native test ότι τα κουμπιά εξακολουθούν να λειτουργούν,
4. μετάβαση DAILY_INPUT / THERAPIST_DAILY και επιστροφή στο MASTER με αυτόματο re-pin.

Το focused source gate έχει ήδη περάσει: **39 passed in 1.71s**.

Μετά συνεχίζει το patient-program gate:

- edit-existing-program end to end,
- optional manual names persist/reload,
- full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Build V74 από clean V48.
2. Native smoke του pinned MASTER toolbar.
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
