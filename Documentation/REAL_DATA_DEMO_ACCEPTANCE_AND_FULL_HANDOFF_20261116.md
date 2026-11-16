# DIOMIL: πραγματικά δεδομένα, παράλληλη αξιολόγηση και πλήρης παράδοση (2026-11-16)

## 2026-11-16 — Executable private real-source DEMO preflight

**New executable interface:** `Python/tools/prepare_real_data_practice_demo.py` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Default operation is **read-only full real-source validation** for an authorized current A:V manual XLSX, confirmed review XLSX, entirely empty application-ready VBA XLSM and independently approved private `clinical_approvals.json`. Uses the strict production source alignment, all-patient approved ID check, original weekday/time/provider reconciliation, unresolved F/U/V stop gates, existing native shell checks and SHA-256 stability of the clinical-approval JSON. Every input must remain off Git checkout. No real patient name, ID, room, doctor, MRN, source path or clinical exception content is emitted to stdout/stderr on handled errors. `--build --output <PRIVATE_*_DEMO_*.xlsm>` is an *explicit* Windows + native Excel/COM, exclusive-new-file second step and repeats the same approval gates. **No real source file has been run; no real-data XLSM exists, and end-to-end practical acceptance is still pending.** Previous synthetic/H8 workbook results do not count as the operator's comparison. First actual action is an authorized private preflight with current manual and approved clinical review, not a random synthetic seed. Native failure or changed approvals after a native operation requires freezing the candidate, not treating it as accepted. See the dedicated operator CLI runbook for exact commands/JSON schema.

**Current blocker:** F/U/V and independently sourced outpatient bookings remain unavailable to the current *inpatient-only* A:V admission bridge. If occupied, the complete real-data source fails closed, never silently drops treatments. Do not alter frozen H2-H10 evidence. Update all living handoff/status/test and delivery documents again with actual source-specific counts and results **only after** the first private native execution, without publishing PII.

**Κατάσταση:** Νέα, ρητή επιχειρησιακή απόφαση. Το **DEMO αξιολόγησης πρέπει να έχει τα πραγματικά δεδομένα του σημερινού χειροκίνητου προγράμματος**. Synthetic/τυχαία δεδομένα επιτρέπονται ΜΟΝΟ ως τεχνικά fixtures/CI και ποτέ ως το τελικό παραδοτέο προς αξιολόγηση. **Δεν έχει παραδοθεί ούτε ελεγχθεί native real-data DEMO.**

**Repository:** `diomil73/EKA-REHAB-TIMETABLE`  
**Branch εργασίας:** `work/afternoon-medical-scope-20261115`  
**Τελευταίο επιβεβαιωμένο γενικό CI:** commit `9fcc610831211dc1baca900d4382771de857d8a3`, run `43769366651`, 1.960 pytest GREEN σε Linux και Windows, synthetic-safety GREEN και στα δύο. Αυτή είναι **source/synthetic regression**, ΟΧΙ Windows Excel COM proof πραγματικού αρχείου ή πρακτική αποδοχή.  
**Πρώτη αρχή:** ΠΡΑΓΜΑΤΙΚΑ δεδομένα σε ιδιωτική, εξουσιοδοτημένη, απομονωμένη δοκιμή. **Χωρίς αλλαγές στο χειροκίνητο πρωτότυπο**.

## 1. Τι παραδίδουμε και τι ΔΕΝ παραδίδουμε

Η τελική πρακτική δοκιμή δεν είναι γενική επίδειξη με ανύπαρκτους ασθενείς. Είναι **μία κανονική εβδομάδα εργασίας ΔΕ–ΠΑ** σε λειτουργικό Diomil XLSM, με την ίδια εικόνα που χρησιμοποιεί σήμερα το τμήμα:
- Πραγματικοί ασθενείς, ακριβής PatientID ταυτοποίηση μέσω εγκεκριμένου ιδιωτικού review, ακριβής κλινική/θάλαμος και υπεύθυνος ιατρός όπου έχουν επαληθευτεί.
- Ακριβής θεραπεία, ημέρα, ώρα, θεραπευτής, πρόγραμμα εσωτερικών/εξωτερικών, εγκεκριμένες εξαιρέσεις και ξεχωριστή ημερήσια πραγματικότητα.
- Σωστή προβολή στο γνωστό MASTER/θεραπευτές/ημέρα και στο αναλυτικό εβδομαδιαίο, όχι μόνο σε άγνωστο τεχνικό φύλλο.
- Ασθενής απών, θεραπευτής απών, αντικατάσταση, εισαγωγή, εξιτήριο/λήξη, προσθήκη θεραπευτή και φοιτητή, Save/Close/Reopen/Restore σε φρέσκο απομονωμένο αντίγραφο.
- Κανένα Σάββατο/Κυριακή ως ημέρα θεραπείας. Κενό πεδίο ημερών σε έγκυρη πηγαία θεραπεία σημαίνει ακριβώς ΔΕ–ΠΑ· έλλειψη ώρας ΔΕΝ δημιουργεί θεραπεία.
- Μόνο εγκεκριμένες πραγματικές εγγραφές. Όχι αυθαίρετο mock όνομα, θεραπεία, ώρα ή πάροχος, ούτε «σιωπηλή» παράλειψη δύσκολων πηγαίων εγγραφών.

**Το υπάρχον synthetic H8/CC fixture δεν καλύπτει αυτό το κριτήριο.** Μπορεί να τρέχει παράλληλα αποκλειστικά ως πακέτο τεστ.

## 2. Ασφάλεια και ακριβής πηγή σύγκρισης

1. Το υφιστάμενο χειροκίνητο αρχείο εξακολουθεί να είναι η **επιχειρησιακά ισχύουσα πηγή**. Λαμβάνουμε **φρέσκο ιδιωτικό read-only snapshot** την ίδια ημερομηνία με την επιλεγμένη εβδομάδα αξιολόγησης. Το παλαιό αρχείο της 10/11/2026 δεν θεωρείται αυτόματα ισοδύναμο του σημερινού.
2. Το πρωτότυπο XLSX, το επιβεβαιωμένο identity/review XLSX, οι εγκεκριμένες κλινικές λεπτομέρειες και το clean macro-enabled shell XLSM παραμένουν **μόνο σε εγκεκριμένο τοπικό Windows χώρο**, με πρόσβαση μόνο από εξουσιοδοτημένα άτομα. Δεν μεταφορτώνονται, συγχρονίζονται ή αποθηκεύονται στο δημόσιο GitHub/CI/διαγνωστικά αρχεία. Χρησιμοποιούμε ιδιωτικά SHA-256, όχι καταγραφές ονομάτων/MRN.
3. Ξεχωριστό, νέο, μη υπάρχον `*_DEMO_*.xlsm`, **εκτός** Git checkout, όχι επέμβαση σε baseline, πραγματικό χειροκίνητο, CC/H10 frozen proof ή υπάρχον παραγωγικό αρχείο. Δεν γίνονται ταυτόχρονες διπλές ενημερώσεις προς το χειροκίνητο σύστημα.
4. Απορρίπτουμε ασυμφωνία ID/κλινικής/πηγαίου κελιού, ελλείπουσες εγκρίσεις, conflicts, σπασμένο schema, ή αλλαγή hash. Δεν μεταφέρουμε ιδιωτικό ιστορικό demo ως πραγματικό ιστορικό θεραπείας.
5. Η νομιμοποίηση και εσωτερική έγκριση επεξεργασίας των πραγματικών δεδομένων είναι απαραίτητη πριν από την εκτέλεση του πιλοτικού. Δεν εξάγουμε φωτογραφίες/συνημμένα ή dump αρχείων με προσωπικές εγγραφές.

## 3. Πλήρης, όχι επιλεκτική, συμφωνία με το χειροκίνητο

| Έλεγχος | Αναμενόμενο |
| --- | --- |
| Πληθυσμός | Κάθε επιβεβαιωμένος ενεργός ασθενής, inpatient/outpatient, μοναδική ταυτότητα, καμία διπλή εγγραφή |
| Κλινική/δωμάτιο/ιατρός | Πιστή αντιστοίχιση ίδιου ασθενούς, χωρίς κληρονομιά προηγούμενης κενής πηγαίας στήλης A |
| Θεραπεία | Ίδια ειδικότητα, πλήθος και κατανομή, χωρίς απώλειες ή διπλομετρήσεις |
| Ημέρα | Ίδια πραγματική ΔΕ–ΠΑ· μηδέν ψευδείς θεραπείες Σ/Κ |
| Ώρα | Ίδια έναρξη, πραγματικές αποκλίσεις καταγράφονται ως blocker |
| Θεραπευτής | Ίδιος επαληθευμένος πάροχος και ίδιοι κοινόχρηστοι περιορισμοί χωρητικότητας |
| Εξωτερικοί | Σωστά καταχωρισμένοι και στο δικό τους πρόγραμμα, χωρίς να γεμίζουν ψευδώς κλίνες |
| Ειδικές στήλες F/U/V | Ανακλινόμενο 30', απογευματινό και πισίνα **δεν χάνονται ούτε εφευρίσκονται**· μπλοκάρουν πλήρη πραγματική εισαγωγή μέχρι να οριστούν/ελεγχθούν |
| Απουσίες/αναπληρώσεις | Ίδια ημερομηνία, ίδιος αρχικός ασθενής/θεραπευτής, σεβασμός slots/δεξιοτήτων, ορατό αποτέλεσμα |
| Εισαγωγή/εξιτήριο | Αξιόπιστη αλλαγή τρέχουσας ημέρας χωρίς διαγραφή βασικού προγράμματος/ιστορικού |
| Save/Restore | Με κλείσιμο και ξανάνοιγμα, καμία απώλεια, ορφανά facts ή αλλαγή παλαιού χειροκίνητου |
| UI | Κουμπιά, κρυμμένες τεχνικές επικεφαλίδες/καρτέλες, χρήσιμη και αναγνώσιμη καθημερινή ροή |

Ο έλεγχος γίνεται **ανά γραμμή και ανά treatment/week-day/time/provider tuple**, όχι μόνο με συνολικά πλήθη ή ένα όμορφο δείγμα οθόνης. Διαφωνίες κρατούνται στον **ιδιωτικό** κατάλογο αποκλίσεων με source-row/cell, υπεύθυνο διόρθωσης και τελικό επανέλεγχο. Δημόσιες αναφορές περιέχουν μόνο συγκεντρωτικούς αριθμούς/κωδικούς ασφάλειας.

## 4. Τι έχει υλοποιηθεί και τι λείπει πραγματικά

### Υπάρχοντα στοιχεία κώδικα (όχι απόδειξη λειτουργικού real-data XLSM)
- `trial_source_weekday_projection.py`, `trial_source_alignment.py`, `trial_demo_source_reconciliation.py`: πηγαίο A:V, ταυτοποίηση ανά αρχική γραμμή, πλήρης αντιπαραβολή tuples, κανόνας ΔΕ–ΠΑ, χωρίς μερική προαγωγή όταν υπάρχουν blockers.
- `trial_patient_registry_projection.py`, `trial_import_plan.py`: εξουσιοδοτημένα πεδία/μοναδικά IDs, source SHA, read-only admission gate.
- `trial_demo_native_seed.py`: Windows COM staged writer PATIENTS / PATIENT_PLANNER / PROGRAM_WEEKLY_SLOTS / DEMO_WEEKLY_VIEW, χωρίς overwrite, readback μέσω production readers, επιπλέον ιδιωτικές φόρμες/κουμπιά.
- `trial_demo_weekly_navigation.py`: πλήρες εβδομαδιαίο και επιστροφή στο MASTER.
- `patient_exit_source.py`, `trial_demo_patient_exit_vba.py`, `trial_demo_active_census.py`: ημερομηνία τελευταίας ενεργής ημέρας, ιδιαίτερη κατάσταση ενεργών ασθενών, συντηρητική αποθήκευση/ανίχνευση χωρίς διαγραφή ιστορικού.
- `master_projection.py` και `master_projection_writer.py`: **προαιρετική** ημερομηνιακή προβολή MASTER, όχι αυτόματη μετά από επιβεβαίωση εξιτηρίου.
- `prepare_daily_input_in_place.py`: διαβάζει το ενιαίο πρόγραμμα και τα PATIENT_EXITS πριν δημιουργήσει συνεδρίες της ημέρας.

### Μη εκπληρωμένα κριτήρια παραδοτέου
1. **Πραγματικό ενημερωμένο manual snapshot / review / εγκεκριμένα κλινικά στοιχεία / κατάλληλο καθαρό XLSM δεν είναι διαθέσιμα στο τρέχον GitHub-only περιβάλλον.** Δεν μεταφέρουμε αληθινές εγγραφές σε GitHub για να «περάσουν τα tests».
2. Οι πηγαίες F/U/V (ανακλινόμενο, απογευματινό, πισίνα) ακόμη **κρατούνται**. Εάν είναι συμπληρωμένες στην πραγματική πηγή, το πλήρες seed ορθά κάνει safety stop, όχι partial «επιτυχία».
3. Δεν υπάρχει επιβεβαιωμένο **ιδιωτικό native Windows Excel COM seed/run** για τα πραγματικά δεδομένα, ούτε operator Save/Close/Reopen πράξεων.
4. Το MASTER δεν ανανεώνει ακόμη αυτόματα το δικό του ενεργό census κατά το in-session «Εξιτήριο / Λήξη», ενώ `DEMO_ACTIVE_CENSUS` ακολουθεί την ημερομηνία.
5. Δεν έχουν ολοκληρωθεί **end-to-end** οι εργασίες εισαγωγής/εξιτηρίου, patient/therapist absences, replacements, student/therapist registry, native restoration και ασφαλές reset trial χωρίς απώλεια δεδομένων.
6. Η τρέχουσα δημιουργία είναι engineering API, όχι πλήρης one-click operator installer. Το πράσινο CI δεν αλλάζει αυτή την κατάσταση.

## 5. Ασφαλής σειρά τελικής πρακτικής παράδοσης

1. Ορίζουμε *ακριβώς* ποιο πρόσφατο χειροκίνητο πρόγραμμα και εβδομάδα θα συγκρίνουμε. Όλοι οι original/review/clinical approvals έλεγχοι τρέχουν **τοπικά και read-only**. Η έγκριση πραγματικών εγγραφών δεν τεκμαίρεται από όνομα/χρώμα Excel.
2. Κλείνουμε/ελέγχουμε F/U/V και όλους τους αποκλεισμούς. Full source alignment και per-slot reconciliation. **Μηδέν σιωπηλές αποκλίσεις**.
3. Εγκεκριμένο, εντελώς καθαρό, πλήρως εξοπλισμένο VBA XLSM shell. Ένας **νέος** private staging στόχος κάθε προσπάθεια. COM write + SHA review + independent native Excel readback, ποτέ openpyxl save ενός κλινικού XLSM.
4. Ο χρήστης ανοίγει το private `*_DEMO_*.xlsm` δίπλα στο ίδιο χειροκίνητο snapshot. Έλεγχος όλων των ασθενών/ημερών/ωρών/θεραπευτών/ειδικοτήτων σε MASTER + αναλυτική εβδομαδιαία + daily therapist. Διορθώνουμε αποκλίσεις πριν από λειτουργικές αλλαγές.
5. Σε ξεχωριστά ανακτήσιμα αντίγραφα δοκιμάζουμε τις υπάρχουσες πραγματικές εργασίες, με πραγματικές εγγραφές αλλά **δοκιμαστικές ενέργειες**: προσθήκη/έξοδος ασθενούς, απουσίες και replacements, θεραπευτής/φοιτητής, rollback/restore και Save/Close/Reopen. Κανένα γεγονός δοκιμής δεν συγχρονίζεται σε παραγωγή.
6. Πλήρης ενημέρωση όλων των **ζωντανών** status/handoff/design/acceptance/runbook αρχείων με αποτελέσματα, SHAs/CI, επιτυχίες, αποτυχίες, unresolved blockers, συγκεκριμένο επόμενο βήμα. Τα παλαιά frozen H2–H10 ιστορικά checkpoints **δεν ξαναγράφονται**: παραπέμπουμε σε αυτά. Μόνο μετά από πραγματική αποδοχή χαρακτηρίζουμε το DEMO «λειτουργικό».

## 6. Εργαλεία ελέγχου και οδηγίες αποτυχιών

Τα παρακάτω εκτελούνται **μόνο στον εξουσιοδοτημένο Windows υπολογιστή** με τα ιδιωτικά αρχεία. Δεν είναι importer ούτε αλλάζουν δεδομένα:

```powershell
python Python\tools\audit_trial_source_alignment.py --programme "<PRIVATE_CURRENT_MANUAL>.xlsx" --review "<PRIVATE_APPROVED_REVIEW>.xlsx"
python Python\tools\preflight_trial_excel_sources.py --review "<PRIVATE_APPROVED_REVIEW>.xlsx" --cc "<PRIVATE_ELIGIBLE_SHELL>.xlsm"
python Python\tools\audit_pilot_source_readiness.py --workbook "<PRIVATE_ELIGIBLE_SHELL>.xlsm"
```

Προσοχή: η CLI `preflight_trial_excel_sources.py` ελέγχει τον **συνήθη populated pilot gate** και μπορεί να εμφανίσει `EMPTY_OR_MISSING_PATIENT_REGISTRY` ακόμη και σε επιτρεπτό κενό DEMO shell. Η συγκεκριμένη `allow_empty_demo_shell=True` εξαίρεση παρέχεται από το **εσωτερικό API** `inspect_native_pilot_inputs` / `prepare_trial_import_plan`, όχι από ανεπιβεβαίωτη επιλογή CLI. Μην ερμηνεύετε τη συνήθη CLI ως πλήρες πράσινο for empty shell.

Αν υπάρχει ένας blocker: **SAFETY STOP**, καμία παραγωγή demo ή overwrite. Κρατήστε ιδιωτικά τον λόγο και τα ελάχιστα διαγνωστικά χωρίς προσωπικά στοιχεία. Αν υπάρχει αποτυχημένο native save, παγώστε το συγκεκριμένο δοκιμαστικό αντίγραφο, μην το «διορθώσετε» ή το ξανατρέξετε στο ίδιο path. Επανεκκίνηση μόνο από νέα καθαρή πηγή/μοναδικό output μετά από τεκμηριωμένη αιτία.

## 7. Live documentation ownership, ενημέρωση και αναφορές

Η **παρούσα** είναι η ενιαία απόφαση για real-data validation. Ενημερώνονται παράλληλα:
- `Documentation/CURRENT_HANDOFF.md` (ακριβής τρέχουσα κατάσταση και συνέχεια).
- `Documentation/PROJECT_STATUS.md` (milestones, blockers, επόμενες πράξεις).
- `Documentation/ONBOARDING.md` (πρώτη ανάγνωση νέας συνεδρίας/συνεργάτη).
- `Documentation/IMPLEMENTATION_MAP.md` (ownership και εκκρεμότητες κώδικα).
- `Documentation/DECISIONS.md` (αμετάκλητος κανόνας real-data trial).
- `Documentation/PROJECT_DELIVERY_MASTER_REPORT_20261115.md` (συγκεντρωτικά παραδοτέα/αποδείξεις).
- `Documentation/INTEGRATED_TEST_BASELINE_MATRIX_20261115.md` (synthetic vs native vs operator acceptance).
- `Documentation/PRACTICE_DEMO_NATIVE_SEED_RUNBOOK_20261115.md` και `Documentation/TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md` (ασφάλεια πραγματικής εισαγωγής).
- `Documentation/REAL_MANUAL_PROGRAM_COLUMNS_AUTHORITATIVE_20261110.md`, `REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md`, `PARALLEL_PILOT_ADMISSION_20261109.md`, `DEMO_LIFECYCLE_AND_ROSTER_REQUIREMENTS_20261112.md`, `SYSTEM_RULES.md` (κλειδωμένοι κανόνες).
- `README.md` και παλαιά resume/continuity entry points ως δείκτες στη νέα απόφαση.

**Ιστορικά `CHECKPOINT_H*` και `RECOVERY_*`: παγωμένα ως αποδείξεις.** Δεν αλλάζουμε παλαιότερα αποτελέσματα, και δεν επαναλαμβάνουμε επιτυχημένες δοκιμές χωρίς νέο σχετικό ρίσκο.

## 8. Τρέχουσα ετυμηγορία

**REAL-DATA DEMO: NOT YET DELIVERED, NATIVE OPERATOR ACCEPTANCE PENDING.** Πρώτη προτεραιότητα η πλήρης, ακριβής, ιδιωτική αντιπαραβολή με το χειροκίνητο πρόγραμμα, όχι άλλο γενικό synthetic demo ή καλλωπισμός στατιστικών. Καμία ανάγκη να γράφει ο χρήστης συνεχώς «οκ» για κάθε μικρό engineering checkpoint.

## AUTHORITATIVE LATEST RESUME 2026-11-16 — REAL-DATA BUSINESS DEMO

**Additional admission blocker made explicit:** The present `project_trial_demo_sources` real A:V weekly projection is *inpatient-only*. Real outpatient appointments need separate, independently approved source/admission via the existing outpatient schedule machinery, with complete cross-specialty/time comparison. Do not claim that the inpatient A:V source already reconstructs all outpatient schedules. If any real outpatient commitments are absent, the whole operator-real-data evaluation is incomplete. The operator must see both types and their genuine appointments. The current `preflight_trial_excel_sources.py` public CLI is a restrictive gate; an otherwise valid empty DEMO shell may return a nonzero code `EMPTY_OR_MISSING_PATIENT_REGISTRY`, which must be handled using the dedicated internal `empty_demo_shell=True` contract rather than bypassing the safety check.


## 2026-11-16 — F/U/V read-only gate, 1,991 tests GREEN

**Concrete new gate, not a synthetic replacement for the user's real-data DEMO:** `Python/rehab_excel/trial_special_fields_audit.py`, `Python/tools/audit_real_manual_specials.py`, `Tests/test_trial_special_fields_audit.py` provide the first **single-source-only** private read-only inspection of actual manual A:V columns **F** (reclining, known 30-minute ΦΘ-related operation), **U** (legacy afternoon, no auto-import) and **V** (morning pool). No review/clinical approval/clean XLSM required for this *diagnostic* first command. It checks original source sheet/header, SHA-256 before/after, cached and raw Excel formula values (never omit formula-only cells), missing-patient source rows, time parseability, all counts and no PII emitted. The tool **never creates sessions, guesses F/V weekdays or imports old U**. It returns 0 if F/U/V truly empty (not whole source approval), 3 for pending real special records, 2 on safety stop. Exact local command, interpretations and privacy gates: `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md`.

**Verified code evidence:** `95157af7dbb32c6cb8aca210c503ffbb7f63be17`, GitHub Actions `43804737124`, **1,991 passing pytest in both Windows and Linux**, all four full/synthetic CI jobs successful. This is not proof that the real private manual has been scanned or a working native Excel was built. **Real-data DEMO remains NOT DELIVERED.** F/V day/provider/duration-aware scheduling approval, U source scope and independently reviewed outpatient programme remain blockers. Next meaningful step is run the read-only audit against the operator-approved private CURRENT manual snapshot, capture only totals, finish explicit F/V mapping and standalone outpatient source, then eligible private native seed and full side-by-side weekday/operator acceptance. Frozen H2–H10 proof artifacts remain untouched.
