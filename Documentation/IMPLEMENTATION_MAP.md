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

### Current focus: V72 native verification

Το V71 πλέον χτίζει και ανοίγει, και το conflict validation φαίνεται λειτουργικά σωστό σε native χρήση. Απέμειναν δύο presentation/consistency follow-ups που έχουν ήδη διορθωθεί σε source/test level:

- το Save conflict dialog ήταν πολύ κοντό και έκρυβε μέρος του μηνύματος και τα κουμπιά **Συνέχεια / Επιστροφή**,
- η ώρα Ρομποτικού της ασθενούς Γιαννιά εμφανιζόταν στο MASTER ως raw decimal Excel time serial.

Η διόρθωση είναι γενική:

- το conflict dialog εφαρμόζει runtime geometry όταν ανοίγει,
- το preview build περνά όλους τους υπάρχοντες ασθενείς από `RefreshMasterPatientProgram()` ώστε τα MASTER treatment cells να ξαναγραφτούν ως composite text (`hh:mm` + ημέρες + όνομα όπου υπάρχει).

Focused gate για αυτά τα δύο fixes: **57 passed in 3.96s**.

Επόμενο βήμα: build και native έλεγχος του `MASTER_TOOLBAR_SMOKE_V72.xlsm` από clean V48.

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
- V71: build/open επιτυχές μετά τα installer hardenings και το conflict flow φαίνεται λειτουργικά σωστό σε native χρήση.
- Focused self-conflict regression gate: **32 passed in 1.40s**.
- Focused V71 conflict/installer gate: **62 passed in 3.34s**.
- Focused V72 prebuild gate για runtime conflict layout + existing MASTER refresh: **57 passed in 3.96s**.

### ΦΘ self-conflict fix

Το same-patient/same-treatment `base_entry_id` γίνεται `ignore_session_id` και αφαιρείται από το effective session set για therapist workload, therapist availability και patient availability. Η συνεδρία που αντικαθίσταται δεν μπλοκάρει τον εαυτό της, ενώ πραγματικές συνεδρίες άλλων ασθενών παραμένουν blocking.

### Conflict-validation batch

Το `validate-vba` στο `patient_program_bridge.py`:

- συγκρίνει όλα τα included complete treatments ανά ζεύγος,
- κοιτά μόνο κοινές weekdays,
- χρησιμοποιεί 45' standard και 30' Ανακλινόμενο,
- επιτρέπει ακριβώς adjacent intervals,
- επιστρέφει explicit diagnostics με ημέρα, θεραπείες και ranges.

Το installed VBA:

- προσθέτει `ValidatePatientProgram()` στο `PatientProgramRuntime`,
- ελέγχει πριν από Αυτόματη συμπλήρωση,
- ελέγχει πριν από οποιοδήποτε Save write,
- δείχνει **Συνέχεια / Επιστροφή** για explicit override.

### V71 installer blockers που έκλεισαν

- VBE `Width/Height` COM rejection στο conflict form: guarded installer geometry.
- brittle `patient-program validation runtime` marker: procedure-scoped insertion μετά το `CompletePatientProgram()`.

### Runtime-layout fix του conflict dialog

Προστέθηκε `patient_program_conflict_layout_vba.py`.

Το conflict form πλέον:

- εφαρμόζει `Me.Width/Height` στο runtime,
- μεγαλώνει το `lblConflict`,
- τοποθετεί ρητά τα κουμπιά **Συνέχεια** και **Επιστροφή**,
- δεν εξαρτάται από το αν το VBE designer δέχθηκε geometry properties κατά το build.

### Initial MASTER composite refresh

Στο `_install_scheduling_ui()` προστέθηκε `_refresh_existing_master_patient_programs(workbook)`.

Μετά την εγκατάσταση/verification του renderer:

- διαβάζει όλα τα nonblank PatientIDs από `PATIENTS`,
- καλεί `RefreshMasterPatientProgram` για κάθε υπάρχοντα ασθενή,
- έτσι τα ήδη υπάρχοντα MASTER treatment cells περνούν από το ίδιο composite formatting path,
- στόχος είναι να εξαφανιστούν stale raw time serials όπως το decimal Ρομποτικό που παρατηρήθηκε για τη Γιαννιά.

Το source/test gate αυτών των δύο τελευταίων fixes είναι πλέον πράσινο: **57 passed in 3.96s**. Native verification εκκρεμεί.

---

## Τρέχον verification gate

Πριν θεωρηθεί κλειστό το V72 batch χρειάζονται:

1. build του `MASTER_TOOLBAR_SMOKE_V72.xlsm` από clean V48,
2. native επιβεβαίωση ότι το conflict dialog δείχνει ολόκληρο το μήνυμα και τα δύο κουμπιά,
3. native επιβεβαίωση ότι η Γιαννιά και άλλοι υπάρχοντες ασθενείς δεν έχουν decimal treatment times στο MASTER,
4. επανέλεγχος Save **Επιστροφή / Συνέχεια**,
5. adjacency test ΦΘ 09:15-10:00 + Ανακλινόμενο 10:00-10:30,
6. edit-existing-program flow,
7. optional manual names persist/reload,
8. full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Build `MASTER_TOOLBAR_SMOKE_V72.xlsm` από clean `MASTER_TOOLBAR_SMOKE_V48.xlsm`.
2. Native smoke: conflict dialog layout + Γιαννιά Ρομποτικό `hh:mm` + spot-check άλλων MASTER cells.
3. Reconfirm conflict Return/Continue και adjacency boundary.
4. Μετά edit-existing-program + optional names.
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
5. active patient-program tests

Σε Windows χρησιμοποιούμε πάντα `python -m pytest`, ποτέ bare `pytest`.
