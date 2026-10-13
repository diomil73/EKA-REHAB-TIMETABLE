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
- Exact adjacency επιτρέπεται.
- Recliner adjacency με ΦΘ: strong soft preference.
- `PATIENT_PLANNER` authoritative, `MASTER_SCHEDULE` derived view.

### Current focus: native V73 verification

Το νέο edit-draft replacement model και τα precise diagnostics έχουν πλέον περάσει και narrow και wider source/integration gates.

Η τρέχουσα υλοποίηση:

- αφαιρεί από το existing-patient baseline όλες τις παλιές εγγραφές treatments που συμμετέχουν στο τρέχον draft,
- λύνει το νέο draft ως ενιαίο σύνολο και ελέγχει τις νέες επιλογές μεταξύ τους,
- κρατά exact adjacency νόμιμη,
- δίνει detailed joint conflict diagnostics με ημέρα, treatments και ranges,
- για manual ΦΘ/provider failure, όταν υπάρχει concrete blocker, εμφανίζει ημέρα, ώρα, θεραπευτή, ασθενή και treatment που κρατά το slot.

Verification checkpoints:

- narrow automatic-scheduling gate: **16 passed in 2.87s**
- wider patient-program integration gate: **71 passed in 3.66s**

Επόμενο βήμα: build του `MASTER_TOOLBAR_SMOKE_V73.xlsm` από clean V48 και native behavior verification.

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

- runtime `Me.Width/Height`,
- explicit sizing του `lblConflict`,
- explicit placement **Συνέχεια / Επιστροφή**,
- ανεξαρτησία από VBE designer geometry persistence.

### Initial MASTER composite refresh

- όλα τα existing PatientIDs περνούν από `RefreshMasterPatientProgram()` κατά το preview build,
- στόχος: να μην παραμένουν stale decimal Excel times σε MASTER cells, όπως παρατηρήθηκε στο Ρομποτικό της Γιαννιά.

### Edit-draft replacement fix

Στο `automatic_patient_schedule.py`:

- δημιουργείται set όλων των treatments του τρέχοντος request/draft,
- persisted assignments του ίδιου ασθενή για αυτά τα treatments αγνοούνται ως values being replaced,
- οι νέες assignments ελέγχονται joint μέσω `_patient_conflict`,
- stale old ΦΘ δεν μπορεί πλέον να μπλοκάρει νέο Ανακλινόμενο όταν και τα δύο συμμετέχουν στο ίδιο edit draft.

### Detailed diagnostics

Προστέθηκαν:

- concrete provider occupancy reason για manual ΦΘ με weekday/time/provider/patient/treatment,
- detailed existing-patient conflict reason με treatment intervals,
- detailed joint-draft conflict reason με shared weekdays και ακριβή ranges.

Regression coverage περιλαμβάνει:

- `Ανακλινόμενο 11:00-11:30` + νέα `ΦΘ 11:30-12:15` με stale old ΦΘ,
- occupied manual ΦΘ slot που ονοματίζει provider και blocking patient,
- πραγματικό joint overlap που εμφανίζει ακριβή ranges.

---

## Τρέχον verification gate

Πριν θεωρηθεί κλειστό το τρέχον batch χρειάζονται:

1. build `MASTER_TOOLBAR_SMOKE_V73.xlsm` από clean V48,
2. native manual ΦΘ occupied-slot test με concrete patient diagnostic,
3. native `Ανακλινόμενο 11:00` + `ΦΘ 11:30` adjacency test,
4. native true-overlap test με explicit treatment/day/time diagnostic,
5. conflict dialog full layout + Γιαννιά Ρομποτικό `hh:mm`,
6. edit-existing-program flow,
7. optional manual names persist/reload,
8. full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Build V73 από clean V48.
2. Native smoke των precise diagnostics, adjacency και presentation fixes.
3. Μετά edit-existing-program + optional names.
4. Full regression suite.

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
