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

### Current focus: precise diagnostics + σωστό edit-draft replacement model

Νέο native smoke έδειξε δύο θέματα:

1. Όταν manual ΦΘ ώρα/provider δεν είναι διαθέσιμα, το μήνυμα είναι πολύ γενικό και δεν λέει ποιος ασθενής καταλαμβάνει το slot.
2. `Ανακλινόμενο 11:00-11:30` με νέα `ΦΘ 11:30-12:15` αναφέρθηκε λανθασμένα ως conflict κατά automatic completion.

Το δεύτερο δεν ήταν duration bug. Η interval λογική ήδη επιτρέπει adjacency. Η αιτία ήταν ότι σε multi-treatment edit ο solver έβλεπε stale persisted assignments των άλλων treatments του ίδιου ασθενή. Δηλαδή το νέο Ανακλινόμενο μπορούσε να συγκρουστεί με την παλιά ΦΘ, παρότι η νέα ΦΘ του draft ήταν 11:30.

Η νέα υλοποίηση:

- αφαιρεί από το existing-patient baseline όλες τις παλιές εγγραφές treatments που συμμετέχουν στο τρέχον draft,
- λύνει το νέο draft ως ενιαίο σύνολο και ελέγχει τις νέες επιλογές μεταξύ τους,
- κρατά exact adjacency νόμιμη,
- δίνει detailed joint conflict diagnostics με ημέρα, treatments και ranges,
- για manual ΦΘ/provider failure, όταν υπάρχει concrete blocker, εμφανίζει ημέρα, ώρα, θεραπευτή, ασθενή και treatment που κρατά το slot.

Το batch είναι implemented και περιμένει focused pytest + native verification.

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
- Focused prebuild gate για runtime conflict layout + existing MASTER refresh: **57 passed in 3.96s**.

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

### Νέο edit-draft replacement fix

Στο `automatic_patient_schedule.py`:

- δημιουργείται set όλων των treatments του τρέχοντος request/draft,
- persisted assignments του ίδιου ασθενή για αυτά τα treatments αγνοούνται ως values being replaced,
- οι νέες assignments ελέγχονται joint μέσω `_patient_conflict`,
- stale old ΦΘ δεν μπορεί πλέον να μπλοκάρει νέο Ανακλινόμενο όταν και τα δύο συμμετέχουν στο ίδιο edit draft.

### Νέα detailed diagnostics

Προστέθηκαν:

- concrete provider occupancy reason για manual ΦΘ με weekday/time/provider/patient/treatment,
- detailed existing-patient conflict reason με treatment intervals,
- detailed joint-draft conflict reason με shared weekdays και ακριβή ranges.

Regression tests προστέθηκαν για:

- `Ανακλινόμενο 11:00-11:30` + νέα `ΦΘ 11:30-12:15` με stale old ΦΘ,
- occupied manual ΦΘ slot που ονοματίζει provider και blocking patient,
- πραγματικό joint overlap που εμφανίζει ακριβή ranges.

Το νέο batch δεν έχει ακόμη user-run test gate.

---

## Τρέχον verification gate

Πριν θεωρηθεί κλειστό το τρέχον batch χρειάζονται:

1. focused `python -m pytest Tests/test_automatic_patient_schedule.py -q`,
2. σχετικό focused patient-program suite αν το πρώτο είναι πράσινο,
3. fresh preview build από clean V48,
4. native manual ΦΘ occupied-slot test με concrete patient diagnostic,
5. native `Ανακλινόμενο 11:00` + `ΦΘ 11:30` adjacency test,
6. native true-overlap test με explicit treatment/day/time diagnostic,
7. conflict dialog full layout + Γιαννιά Ρομποτικό `hh:mm`,
8. edit-existing-program flow,
9. optional manual names persist/reload,
10. full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Pull latest branch.
2. Run focused automatic scheduling tests.
3. Αν είναι πράσινα, rebuild επόμενου preview από clean V48.
4. Native smoke των δύο νέων observations και των presentation fixes.
5. Μετά edit-existing-program + optional names.
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
