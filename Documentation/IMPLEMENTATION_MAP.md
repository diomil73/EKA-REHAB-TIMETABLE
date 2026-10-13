# Implementation Map

Τελευταία ενημέρωση: 2026-10-13

## Σκοπός

Ο ζωντανός χάρτης υλοποίησης του project. Για κάθε ουσιαστικό milestone καταγράφει: **τι φτιάχνουμε τώρα**, **τι έχει ολοκληρωθεί**, **ποιο είναι το verification gate** και **ποια είναι τα επόμενα βήματα**.

Για το ακριβές σημείο συνέχισης μετά από pause/context loss, πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

---

## Τι φτιάχνουμε τώρα

### Patient-centric multi-treatment scheduling

Μία ενιαία patient-program φόρμα για νέο και υπάρχοντα ασθενή, με πολλές θεραπείες και ένα τελικό atomic save.

Κύριοι κανόνες:

- ΦΘ: provider υποχρεωτικός και solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ: ημέρες + ώρα, optional free-text όνομα, όχι solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ: χωρίς provider/name.
- Manual nonblank day/time/provider constraints δεν αντικαθίστανται σιωπηρά.
- Standard treatment duration: 45 λεπτά.
- Ανακλινόμενο: 30 λεπτά.
- Recliner adjacency με ΦΘ: strong soft preference.
- `PATIENT_PLANNER` authoritative, `MASTER_SCHEDULE` derived view.

### Current focus: native verification του κοινού conflict gate

Το V70 native smoke έδειξε ότι μια χειροκίνητη σύγκρουση, π.χ. ΦΘ και Λογο την ίδια Δευτέρα σε overlapping ώρα, εντοπιζόταν από την **Αυτόματη συμπλήρωση** αλλά μπορούσε να αποθηκευτεί κανονικά από το κουμπί **Αποθήκευση**.

Η υλοποίηση του κοινού conflict gate ολοκληρώθηκε και το focused regression gate πέρασε. Το πρώτο V71 build όμως σταμάτησε πριν το save κατά την εγκατάσταση του νέου conflict dialog λόγω VBE COM property error στο `Width`.

Η ζητούμενη συμπεριφορά παραμένει:

- αναλυτικό μήνυμα, π.χ. `Δευτέρα: ΦΘ 09:00-09:45 συγκρούεται με Λογο 09:00-09:45.`,
- στην Αυτόματη συμπλήρωση η σύγκρουση σταματά τη διαδικασία και ζητά διόρθωση,
- στην Αποθήκευση εμφανίζεται modal επιλογή **Συνέχεια** / **Επιστροφή**,
- **Επιστροφή** σημαίνει μηδενικά writes,
- **Συνέχεια** είναι ρητή user override και επιτρέπει intentional conflicting save.

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
- Atomic final save με snapshots, verification, rollback και single workbook save.
- CURRENT preload από `PATIENT_PLANNER` για edit υπάρχοντος προγράμματος.
- Ίδια patient-program φόρμα μετά από **Επεξεργασία ασθενή**.
- Optional manual names για Ρομποτικό / Εργο / Λογο / Ψυχ.
- Duration-aware patient conflicts.
- ΦΘ / Ανακλινόμενο adjacency scoring.
- MASTER composite renderer: ώρα + ημέρες + provider/name όπου υπάρχει.
- MASTER refresh μετά από patient-program save και registry edit.

### Installer/runtime hardening

- legacy `PatientSchedulingRuntime` και old single-treatment markers απορρίπτονται από read-back verifier.
- CRLF/LF-safe VBE patching.
- procedure-scoped patient-edit refresh patch.
- runtime stage diagnostics για MASTER refresh.
- empty MASTER patient cells guarded πριν από `Split`.
- conflict-form geometry setters πλέον είναι best-effort απέναντι σε VBE COM rejection για `Width/Height` και `Left/Top`.

### Verified native checkpoints

- V69: first save επιτυχές και MASTER composite rendering επιτυχές.
- V70: το προηγούμενο false alarm της αμετάβλητης ΦΘ κατά edit revalidation διορθώθηκε native.
- Focused self-conflict regression gate: **32 passed in 1.40s**.

### ΦΘ self-conflict fix

Το same-patient/same-treatment `base_entry_id` γίνεται `ignore_session_id` και αφαιρείται από το effective session set για:

- therapist workload,
- therapist availability,
- patient availability.

Έτσι η παλιά ΦΘ που αντικαθίσταται δεν μπλοκάρει τον εαυτό της, ενώ πραγματικές συνεδρίες άλλων ασθενών παραμένουν blocking.

### Conflict-validation batch

Προστέθηκε `validate-vba` στο `patient_program_bridge.py`.

Ο validator:

- συγκρίνει όλα τα included complete treatments ανά ζεύγος,
- κοιτά μόνο κοινές weekdays,
- χρησιμοποιεί 45' standard και 30' Ανακλινόμενο,
- επιτρέπει ακριβώς adjacent intervals,
- επιστρέφει explicit `CONFLICT` rows με weekday, treatments, start times και human-readable message.

Το installed VBA πλέον:

- προσθέτει `ValidatePatientProgram()` στο `PatientProgramRuntime`,
- καλεί τον validator πριν από Αυτόματη συμπλήρωση,
- καλεί τον ίδιο validator στην `SavePatientProgramAtomic()` **πριν από worksheet access/snapshots/writes**,
- εγκαθιστά `frmPatientProgramConflict` με κουμπιά **Συνέχεια** και **Επιστροφή**,
- με **Επιστροφή** βγαίνει από το save πριν γραφτεί οτιδήποτε,
- με **Συνέχεια** επιτρέπεται explicit override.

Focused regression gate μετά από αυτό το batch:

**59 passed in 5.89s**

### V71 conflict-dialog installer blocker

Το πρώτο V71 build σταμάτησε με:

`Property '<unknown>.Width' can not be set.`

Η αιτία ήταν ότι ο νέος `frmPatientProgramConflict` installer έγραφε απευθείας `designer.Width/Height` και control `Width/Height`, ενώ οι παλιότεροι installers ήδη αντιμετώπιζαν τέτοιες VBE COM properties ως best-effort.

Το hardening πλέον:

- κάνει guarded το form Width/Height,
- κάνει guarded το control Left/Top,
- κάνει ξεχωριστά guarded το control Width/Height,
- έχει regression test με fake control που απορρίπτει Width/Height.

Το hardening είναι implemented αλλά περιμένει focused pytest και νέο V71 build.

---

## Τρέχον verification gate

Πριν θεωρηθεί verified το conflict batch χρειάζονται:

1. focused pytest μετά το V71 Width/Height hardening,
2. clean rebuild του V71 από V48,
3. native Auto test με ΦΘ + Λογο overlap και explicit μήνυμα ημέρας/ωρών,
4. native Save test με **Επιστροφή** και επιβεβαίωση ότι `PATIENT_PLANNER`/MASTER δεν άλλαξαν,
5. native Save test με **Συνέχεια** και επιβεβαίωση intentional override,
6. adjacency test που επιβεβαιώνει ότι ΦΘ 09:15-10:00 + Ανακλινόμενο 10:00-10:30 δεν θεωρείται conflict,
7. edit-existing-program flow,
8. optional manual names persist/reload,
9. full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. `git pull origin feature/outpatient-scheduling`.
2. Τρέξιμο focused tests με το νέο conflict-form sizing regression.
3. Rebuild `MASTER_TOOLBAR_SMOKE_V71.xlsm` από clean `MASTER_TOOLBAR_SMOKE_V48.xlsm`.
4. Native smoke του conflict flow Auto + Save Return + Save Continue.
5. ΦΘ/Ανακλινόμενο adjacency boundary smoke.
6. Μετά edit-existing-program + optional names.
7. Full regression suite.

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
5. active patient-program tests

Σε Windows χρησιμοποιούμε πάντα `python -m pytest`, ποτέ bare `pytest`.
