# Implementation Map

Τελευταία ενημέρωση: 2026-10-14

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
- Exact adjacency επιτρέπεται.
- Recliner adjacency με ΦΘ: strong soft preference.
- `PATIENT_PLANNER` authoritative, `MASTER_SCHEDULE` derived view.

### Current focus: edit-existing-program flow

Το V73 scheduling/conflict batch έχει πλέον επιβεβαιωθεί native ως πράσινο.

Επόμενο milestone είναι να κλείσει end-to-end το edit υπάρχοντος ασθενή:

- Basic menu → **Επεξεργασία ασθενή**,
- αποθήκευση registry αλλαγών,
- άνοιγμα της ίδιας patient-program φόρμας για το ίδιο PatientID,
- preload του υπάρχοντος προγράμματος από `PATIENT_PLANNER`,
- αλλαγή μίας ή περισσότερων θεραπειών,
- ένα τελικό atomic save,
- σωστό MASTER composite refresh μετά την αλλαγή.

Μετά από αυτό ακολουθεί verification των optional manual names και τέλος full regression.

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
- conflict-form designer geometry setters είναι best-effort απέναντι σε VBE COM rejection.
- patient-program runtime validation patch είναι procedure-scoped και idempotent.

### Verified native checkpoints

- V69: first save επιτυχές και MASTER composite rendering επιτυχές.
- V70: το false alarm της αμετάβλητης ΦΘ κατά edit revalidation διορθώθηκε native.
- V71: build/open επιτυχές μετά τα installer hardenings και το conflict flow λειτουργεί native.
- V73: precise ΦΘ blocker diagnostics, exact adjacency, detailed true-overlap diagnostics, conflict-dialog layout και normal Ρομποτικό time rendering επιβεβαιώθηκαν native.

### Source/integration checkpoints

- Focused self-conflict regression gate: **32 passed in 1.40s**.
- Focused V71 conflict/installer gate: **62 passed in 3.34s**.
- Focused prebuild gate για runtime conflict layout + existing MASTER refresh: **57 passed in 3.96s**.
- Narrow edit-draft/diagnostics gate: **16 passed in 2.87s**.
- Wider patient-program integration gate: **71 passed in 3.66s**.

### Conflict-validation / Save override

Το `validate-vba`:

- συγκρίνει όλα τα included complete treatments ανά ζεύγος,
- κοιτά μόνο κοινές weekdays,
- χρησιμοποιεί 45' standard και 30' Ανακλινόμενο,
- επιτρέπει exactly adjacent intervals,
- επιστρέφει explicit diagnostics με ημέρα, θεραπείες και ranges.

Το installed VBA:

- ελέγχει πριν από Αυτόματη συμπλήρωση,
- ελέγχει πριν από οποιοδήποτε Save write,
- δείχνει **Συνέχεια / Επιστροφή** για explicit override.

### Runtime-layout fix του conflict dialog

Native V73 verified:

- runtime `Me.Width/Height`,
- explicit sizing του `lblConflict`,
- explicit placement **Συνέχεια / Επιστροφή**,
- ανεξαρτησία από VBE designer geometry persistence.

### MASTER composite rendering

- existing PatientIDs περνούν από `RefreshMasterPatientProgram()` κατά το preview build,
- normal patient-program saves/edit saves γράφουν composite treatment cells,
- το προηγούμενο decimal Ρομποτικό της Γιαννιά δεν παραμένει ως λειτουργικό blocker μετά από re-entry/save.

### Edit-draft replacement fix

Στο `automatic_patient_schedule.py`:

- persisted assignments του ίδιου ασθενή για treatments που συμμετέχουν στο τρέχον draft αγνοούνται ως values being replaced,
- οι νέες assignments ελέγχονται joint μέσω `_patient_conflict`,
- stale old ΦΘ δεν μπλοκάρει πλέον νέο Ανακλινόμενο όταν και τα δύο συμμετέχουν στο ίδιο edit draft.

Native V73 confirmed το key scenario `Ανακλινόμενο 11:00-11:30` + `ΦΘ 11:30-12:15`.

### Detailed diagnostics

Προστέθηκαν και επιβεβαιώθηκαν native στο V73:

- concrete provider occupancy reason για manual ΦΘ με weekday/time/provider/patient/treatment,
- detailed existing-patient conflict reason με treatment intervals,
- detailed joint-draft conflict reason με shared weekdays και ακριβή ranges.

---

## Τρέχον verification gate

Πριν θεωρηθεί ολοκληρωμένο το patient-program milestone χρειάζονται:

1. native edit-existing-program flow end to end,
2. optional manual names persist/reload για Ρομποτικό / Εργο / Λογο / Ψυχ,
3. τελικό duration/recliner smoke όπου χρειάζεται,
4. full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Native test υπάρχοντος ασθενή μέσω **Επεξεργασία ασθενή**.
2. Επιβεβαίωση preload του υπάρχοντος program.
3. Αλλαγή treatment δεδομένων και atomic save.
4. Έλεγχος `PATIENT_PLANNER` και MASTER μετά το edit.
5. Optional manual names persist/reload.
6. Full regression suite.

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
