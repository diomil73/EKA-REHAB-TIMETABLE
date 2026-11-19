# DIOMIL | ΕΝΙΑΙΟΣ ΟΔΗΓΟΣ ΑΚΡΙΒΟΥΣ ΣΥΝΕΧΕΙΑΣ

> **ΕΝΕΡΓΗ ΣΥΝΕΧΕΙΑ 17/11/2026:** Ιδιωτικό R4: 846 εβδομαδιαίες προτάσεις, 818 χωρίς σήμανση patient conflict, 28 HOLD_CONFLICT, 29 εκκρεμή ωράρια Ψυχολογίας, 106 μη επιβεβαιωμένες ταυτότητες. Οι έξι ασάφειες διευκρινίστηκαν. Βλέπε §4.36. Native DEMO PENDING.

> **ΤΕΛΕΥΤΑΙΟ VERIFIED CLOUD CHECKPOINT 17/11/2026:** GitHub Actions [43980917796](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43980917796), code SHA `14d095367845cbf0c60aacf84d63c07bd303e899`: **2.040/2.040 Python pytest Windows και 2.040/2.040 Linux, 4/4 jobs SUCCESS**. Δύο νέες μικρές ενισχύσεις της πηγαίας ασφάλειας (6 fictional-only tests συνολικά), §4.7. Αφορά GitHub cloud code, **όχι** το PC, native XLSM ή κλινική αποδοχή. Η ρητή παύση προσωπικού PC στην §4.6 παραμένει σε ισχύ.


> **ΜΟΝΑΔΙΚΟ ΚΕΝΤΡΙΚΟ ΑΡΧΕΙΟ.** Αν ο χρήστης γράψει «Βρες το αρχείο-οδηγό DIOMIL και συνέχισε ακριβώς εκεί που μείναμε», βρες αυτό το αρχείο στον **τρέχοντα** κλάδο GitHub, διάβασέ το ολόκληρο και συνέχισε από το **Κεφάλαιο 4**. Μη βασιστείς σε παλιές περιλήψεις ή ζητήσεις επανάληψη όσων ήδη αποδείχθηκαν. Τα παλιά checkpoints είναι ιστορικά, όχι ανταγωνιστικές οδηγίες.
>
> **Τέσσερα κεφάλαια μέσα στο ΙΔΙΟ αρχείο:** 1. Πού μείναμε και πραγματικά αποδεικτικά, 2. Επιχειρησιακοί κανόνες και υπάρχουσες αποδεκτές λειτουργίες, 3. Αρχιτεκτονική/ασφαλής μέθοδος, 4. Ακριβή επόμενα βήματα, κριτήρια demo και ενημέρωση οδηγού.
>
> **Snapshot σύνταξης: 17/11/2026.** Επαληθευμένο HEAD *πριν* τον παρόντα οδηγό: `38250856ad2e98f70d1aa64253b449422efef250`. GitHub Actions [43971761025](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43971761025): **4/4 success**. Το νέο commit του οδηγού θα έχει νεότερο SHA. Σε κάθε επανεκκίνηση ανακτούμε **live HEAD**, δεν αντιμετωπίζουμε αυτόν τον αριθμό ως μόνιμα τελευταίο.
>
> **Διαβάζουμε προσεκτικά τους ισχυρισμούς.** Source/CI GREEN ≠ πραγματικό native XLSM ≠ κλινική έγκριση ≠ operator acceptance. Ποτέ ονόματα/IDs/στοιχεία ασθενών στο δημόσιο GitHub.

---

# ΚΕΦΑΛΑΙΟ 1. ΤΡΕΧΟΝ ΣΗΜΕΙΟ, ΠΡΑΓΜΑΤΙΚΑ ΑΡΧΕΙΑ, ΚΑΤΑΣΤΑΣΗ ΚΑΙ ΑΠΟΔΕΙΞΕΙΣ

## 1.1 Ποιο έργο, ποιος κλάδος, ποιος στόχος

- Repository: [`diomil73/EKA-REHAB-TIMETABLE`](https://github.com/diomil73/EKA-REHAB-TIMETABLE).
- Ενεργός authoritative branch: **`work/afternoon-medical-scope-20261115`**. Όχι αυτόματη επιστροφή σε `main`, `feature/outpatient-scheduling` ή `work/history-statistics-rebuild`, που εμφανίζονται σε ιστορικές αναφορές.
- Στόχος χρήστη: **ένα χωριστό, ιδιωτικό, πλήρως δοκιμαζόμενο Excel/VBA DEMO που χρησιμοποιεί πραγματικούς ασθενείς και ακριβές εγκεκριμένο τρέχον πρόγραμμα**, για καθημερινή πρακτική αξιολόγηση. Το υπάρχον παραγωγικό χειροκίνητο πρόγραμμα δεν αντικαθίσταται/τροποποιείται. Synthetic fixtures επιτρέπονται για τεχνικά tests, **όχι για αξιολόγηση από τον χρήστη**.
- Εργαζόμαστε αυτόνομα, χωρίς να ρωτάμε «οκ;» μετά από κάθε commit/command. Σταματάμε μόνο για πραγματική ανθρώπινη κλινική απόφαση, μη προσβάσιμο ιδιωτικό αρχείο ή μη διαθέσιμο Windows Excel/COM. Κάθε ουσιαστικό βήμα καταγράφεται στον παρόντα **μοναδικό** οδηγό.
- Οι παλαιές `Documentation/CURRENT_HANDOFF.md`, `PROJECT_STATUS.md`, `ONBOARDING.md` παραμένουν ζωντανές *θύρες* αλλά **αυτό το αρχείο** είναι η πρώτη ανάγνωση και κατέχει την ακριβή ενιαία συνέχεια. Το [PROJECT_DELIVERY_MASTER_REPORT_20261115.md](PROJECT_DELIVERY_MASTER_REPORT_20261115.md) διατηρεί αναλυτικό ιστορικό και αποδεικτικά, όχι ξεχωριστό start point.

## 1.2 Πραγματικό checkpoint, όχι υποθέσεις

| Έλεγχος | Τεκμηριωμένο αποτέλεσμα |
|---|---|
| Τεχνική πηγή | Commit [3825085](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/38250856ad2e98f70d1aa64253b449422efef250), Actions [43971761025](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43971761025), **4/4 GREEN** |
| Πλήρη Python tests | [Actions 43971456352](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43971456352): **2.031 pytest Windows και 2.031 Linux**, fictional/synthetic source regression |
| Πραγματικά XLSX | Τα δύο συνημμένα της προηγούμενης συνομιλίας **ανοίχτηκαν πραγματικά μόνο για ανάγνωση**. SHA πριν/μετά unchanged. Δεν σημαίνει ότι τα ίδια bytes θα είναι σε κάθε μελλοντικό runtime |
| Πρόγραμμα / review | **106/106** patient source-row/name/context reviews ταιριάζουν, **764** προτεινόμενες ΔΕ–ΠΑ εβδομαδιαίες εγγραφές, **0** source-cell provenance mismatches στον ανεξάρτητο έλεγχο |
| PatientID | **0/106 εγκεκριμένα**, **33** παλιά IDs ως *μη εγκεκριμένες* προτάσεις |
| Εκκρεμότητες review | **61** εγγραφές προς έλεγχο, **11** καταγεγραμμένες επικαλύψεις αλλά **15** conflict groups από πλήρη ανεξάρτητο scan: 11 cross-specialty + **4 επιπλέον διπλές Λογοθεραπείες με ίδιο source cell/day/time** |
| F / U / V | **5 F** (30' ανακλινόμενο), **19 U** (ιστορικό απογευματινό), **0 V** (πισίνα δεν εμφανίζεται στο συγκεκριμένο snapshot) |
| Outpatients | Βοηθητικό φύλλο `Εξωτερικοί Ασθενείς` με **10 συμπληρωμένες σειρές** αλλά **όχι** εγκεκριμένη schema πηγή `PATIENTS` + `OUTPATIENT_SCHEDULE`· δεν αποδεικνύει 10 συνεδρίες |
| Native πραγματικό demo | **ΔΕΝ ΕΧΕΙ ΔΗΜΙΟΥΡΓΗΘΕΙ** και δεν έχει γίνει operator acceptance. Δεν έχει γίνει τοπικό real-data COM seed ούτε πλήρης Save/Close/Reopen |
| Προστασία | Original manual/review, παγωμένες H2-H10 αποδείξεις, immutable `Rehab_Center_System_v27_1.xlsm`, existing MASTER renderer: αμετάβλητα |

Τα 764 proposed entries ανά ειδικότητα: **ΦΘ 471, Εργοθεραπεία 154, Λογοθεραπεία 34, Ψυχολογία 34, ΕΦΑ 29, Ρομποτική 25, Ανακλινόμενο 17**. Υπάρχουν 108 source B text rows, αλλά **2** δεν είναι review ασθενείς και δεν περιέχουν C:V treatments· **30** context cells στον πίνακα review λείπουν. Οι 17 εβδομαδιαίες υποψήφιες για ανακλινόμενο **δεν αποτελούν εγκεκριμένη μοντελοποίηση** των 5 αρχικών F.

## 1.3 Ονόματα των πραγματικών συνημμένων, προέλευση και πρόσβαση

- ΠΡΑΓΜΑΤΙΚΟ manual filename: **`ΠΡΟΓΡΑΜΜΑΤΑ  ΘΕΡΑΠΕΙΩΝ ΑΣΘΕΝΩΝ 10-11-26 .xlsx`**. Υπάρχουν **δύο κενά** μεταξύ των δύο πρώτων λέξεων και **ένα κενό πριν από `.xlsx`**. Αυτό εξήγησε το `MANUAL_NAME_MATCH=False` και `MANUAL_FILES_FOUND=0` του ακριβούς τοπικού filename filter. Το αρχείο ήταν σωστό· το όνομα που είχαμε γράψει αρχικά ήταν ανεπαίσθητα διαφορετικό.
- Identity proposal: **`DIOMIL_ΕΛΕΓΧΟΣ_ΤΑΥΤΟΤΗΤΩΝ_ΠΡΟΤΑΣΕΙΣ_20261112.xlsx`**. `ΠΡΟΤΑΣΕΙΣ` **δεν** σημαίνει `review_confirmed`.
- Ημερομηνία 10/11/2026 είναι **ιστορικό snapshot**. Χρειάζεται επαλήθευση ή νέο εγκεκριμένο current manual για την εβδομάδα πραγματικού demo.
- **Αναλυτικό, προστατευμένο από στοιχεία ασθενών τεκμήριο:** [`Documentation/REAL_SOURCE_UPLOADED_PAIR_READONLY_AUDIT_20261117.md`](REAL_SOURCE_UPLOADED_PAIR_READONLY_AUDIT_20261117.md). Προέλευση raw counts, ειδικών, exclusions και συμφωνίας. Δεν δημοσιεύουμε τα αρχικά XLSX/PatientIDs/private paths/hashes.
- Τα actual attachments είχαν ανοίξει σε προηγούμενο conversation/runtime. Αν τώρα δεν είναι διαθέσιμα, βρες πρώτα τα ίδια conversation/Project/Library attachments μόνο όταν χρειάζονται για **νέα** δουλειά. Μη λες «δεν διαβάστηκαν ποτέ» και μην επαναλαμβάνεις το ίδιο source audit από συνήθεια.

## 1.4 Γιατί εμφανίστηκε άσκοπη επανάληψη στο παρελθόν

- Ο τοπικός Windows Git checkout ήταν **155 commits πίσω**. Ο χρήστης ολοκλήρωσε με `git fetch origin` και `git merge --ff-only` έως `3b47567`. Τα δύο audit scripts επαληθεύτηκαν `Test-Path=True`. Το μεμονωμένο `else` error ήταν σφάλμα τμηματικής PowerShell επικόλλησης και δεν εμπόδισε το fast-forward.
- Η επόμενη αναζήτηση με λανθασμένο ακριβές όνομα βρήκε 0 manual και 1 proposal. Ο `SAFETY STOP` πέτυχε, αλλά έπειτα ο χρήστης εκτέλεσε επιπλέον γραμμές με **κενό `$p`**. Τα `argparse: --programme expected one argument` **δεν** είναι αποτυχίες κλινικών δεδομένων, επειδή κανένα script δεν άνοιξε πηγή.
- **Μετέπειτα** ο χρήστης ανέβασε τα δύο XLSX και ο ανεξάρτητος read-only έλεγχος ολοκληρώθηκε. Δεν χρειάζεται ξανά η ίδια διάγνωση.
- Εντοπίστηκαν **4 πραγματικά duplicate same-source Λογοθεραπεία groups** πέρα από τις 11 cross-specialty conflicts. Μπήκαν **2 fictional-only regression tests**: `Tests/test_trial_weekly_projection.py::test_same_speech_therapy_source_slot_cannot_be_imported_twice` και `Tests/test_trial_admission_gate.py::test_identical_speech_therapy_proposals_fail_without_overlap_sheet_rows`. Το ίδιο το πραγματικό review **δεν** διορθώθηκε.

## 1.5 Κρίσιμα commits και στάθμη απόδειξης

1. [`58ceeec`](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/58ceeec5f5a019e27bb9701145378dbb9d7c400f): ανεξάρτητος έλεγχος αληθινών ανεβασμένων αρχείων, read-only report.
2. [`3272528`](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/327252826541bfc0dc8863d280672844d100f037): 4 επιπλέον source-duplicate Λογοθεραπείας.
3. [`6669489`](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/6669489d28945808fcac38f288c8950e7fb55b40): 2 new synthetic safety tests και 2.031 cross-platform green.
4. [`3825085`](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/38250856ad2e98f70d1aa64253b449422efef250): παραδοτέο real-read-only checkpoint και πράσινο 4/4 CI.
5. Το **commit του παρόντος οδηγού** είναι νεότερο. Σε κάθε resume επαλήθευσε live `HEAD` πριν συμπεράνεις κατάσταση.

**Τέσσερις διαφορετικές ενδείξεις, μην τις αναμειγνύεις:** (α) `SOURCE/CI GREEN`, (β) `NATIVE ACCEPTED` μόνο σε συγκεκριμένο frozen preview, (γ) `REAL SOURCE READ-ONLY GREEN` για provenance χωρίς approval, (δ) `REAL NATIVE DEMO ACCEPTED` μετά από current private data, Windows Excel/COM, 1:1 ΔΕ–ΠΑ, workflows, Save/Close/Reopen και υπογραφή χειριστή. **Το (δ) παραμένει PENDING**.

---


# ΚΕΦΑΛΑΙΟ 2. ΚΛΕΙΔΩΜΕΝΟΙ ΚΑΝΟΝΕΣ, ΥΠΑΡΧΟΥΣΕΣ ΛΕΙΤΟΥΡΓΙΕΣ ΚΑΙ ΟΡΙΑ

## 2.1 Πρόγραμμα, ημέρες, βασικό επιχειρησιακό μοντέλο

- Η θεραπευτική εβδομάδα είναι **μόνο Δευτέρα έως Παρασκευή**. Δεν υπάρχουν ραντεβού/θεραπείες Σαββάτου–Κυριακής. Αυτός ο κανόνας είναι ρητά επιβεβαιωμένος από τον χρήστη.
- Το χειροκίνητο αρχικό πρόγραμμα είναι η βάση του **recurring εβδομαδιαίου σχεδίου** και δεν αντικαθίσταται από ημερήσιες εξαιρέσεις. Απουσία ασθενούς, απουσία θεραπευτή, ακύρωση, αντικατάσταση είναι **ημερήσιο operational overlay**. Δεν καταστρέφουν ούτε μεταλλάσσουν τα αρχικά recurring assignments.
- Στο **A:V source import** ειδικά, όταν υπάρχει **ώρα** θεραπείας αλλά το αντίστοιχο πεδίο ημερών είναι **κενό**, θεωρούμε **ΔΕ–ΠΑ** ακριβώς. Αυτό είναι local source-import `implicit_daily=True`· **όχι** γενική αλλαγή default όλων των parsers. Όταν **δεν υπάρχει ώρα**, δεν δημιουργείται slot. Ημέρες/provider/κείμενο χωρίς ώρα απαιτούν ξεχωριστό review.
- Οι διαφορετικές ώρες στις διαφορετικές μέρες πρέπει να διατηρούνται ως διακριτές πηγαίες εβδομαδιαίες τοποθετήσεις. Απαγορεύεται να μετατραπούν πολλές πραγματικές ώρες σε μία γενική σταθερή ώρα ή να εξαφανιστεί η πληροφορία του αρχικού source cell.
- **Πρωτότυπο πρόγραμμα `ΠΡΟΓΡΑΜΜΑ` A:V:**

| Στήλη | Πηγαία σημασία | Στη νέα εισαγωγή |
|---|---|---|
| A | Κλινική / πηγαίο context | Δεν συμπληρώνεται από την προηγούμενη σειρά χωρίς ανεξάρτητη έγκριση |
| B | Ονοματεπώνυμο ασθενούς | Όχι authoritative PatientID |
| C, D, E | Φυσικοθεραπεία: ώρα, ημέρες, θεραπευτής | Ακριβής ΦΘ recurrence/provider |
| F | Ανακλινόμενο, ώρα | ΦΘ-σχετικό **30 λεπτών**, μη αυτόματη αυθαίρετη ημέρα/provider |
| G, H, I | Εργοθεραπεία: ώρα, ημέρες, θεραπευτής | Ακριβής διατήρηση ημερών/ωρών |
| J, K, L | Λογοθεραπεία: ώρα, ημέρες, θεραπευτής | Διπλά ίδια source-cell slots απορρίπτονται |
| M, N, O | Ψυχολογία: ώρα, ημέρες, θεραπευτής | Ακριβής διατήρηση |
| P, Q | ΕΦΑ: ώρες, ημέρες | Δεσμεύσεις μόνο με πραγματική ώρα |
| R, S, T | Ρομποτική: ώρα, ημέρες, θεραπευτής | Μόνο εγκεκριμένοι πάροχοι/ικανότητες |
| U | Παλαιό απογευματινό | Ιστορικό **όχι** αυτόματο νέο πρωινό/απογευματινό |
| V | Πισίνα | Πρωινή ειδικότητα, μόνο με εγκεκριμένα στοιχεία |

- **F/U/V source gate:** Το F χρειάζεται εγκεκριμένη συσχέτιση ημερών/provider/capacity και σύνδεση με ΦΘ, το U ρητή πολιτική για ιστορικό απόγευμα, το V πλήρη ημέρα/ώρα/provider. Η σημερινή `prepare_real_data_practice_demo.py` κάνει **hard stop** όταν υπάρχει οποιοδήποτε F/U/V. Άρα ακόμη και μια ανθρώπινη έγκριση δεν αρκεί χωρίς κατάλληλη verified εισαγωγική λειτουργία ή εγκεκριμένη ρητή διαδικασία. Απαγορεύεται σιωπηλή παράλειψη/σβήσιμο πεδίων. **V=0** σε αυτό το snapshot **δεν** σημαίνει κατάργηση της πισίνας από το Diomil.
- Πηγή λεπτομερειών: [TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md](TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md), [REAL_MANUAL_PROGRAM_COLUMNS_AUTHORITATIVE_20261110.md](REAL_MANUAL_PROGRAM_COLUMNS_AUTHORITATIVE_20261110.md), [SYSTEM_RULES.md](SYSTEM_RULES.md), [DECISIONS.md](DECISIONS.md).

## 2.2 Ασθενείς, ταυτότητες, εξωτερικοί και capacity

- `PatientID`: **αμετάβλητο εσωτερικό identity key**, δεν μεταβάλλεται όταν αλλάζει κλινική/θάλαμος/ιατρός/θεραπευτής. Ούτε όνομα ούτε παλιό προτεινόμενο ID αρκεί για αυτόματη εξουσιοδότηση. `HospitalMRN`: ξεχωριστό προαιρετικό νοσοκομειακό αναγνωριστικό. Κάθε πραγματική καταχώριση/αντιστοίχιση επιβεβαιώνεται ιδιωτικά από εξουσιοδοτημένο χειριστή.
- Εσωτερικοί και εξωτερικοί είναι πλήρεις κατηγορίες ασθενών. Ο εξωτερικός δεσμεύει πραγματικό capacity/therapist workload και συμμετέχει σε daily replacements, ιστορικό και scheduling. Δεν είναι dummy ή μόνο διακοσμητικό κείμενο.
- Ο ίδιος ακριβής ασθενής/ημέρα/ώρα δεν μπορεί να κρατά δύο ασύμβατες θεραπείες, ενώ το ίδιο source-cell/day/time δεν δημιουργεί δύο πραγματικές συνεδρίες. Shared therapist slots θέλουν **ξεχωριστή εγκεκριμένη ακριβή membership** όλων των PatientIDs· δεν επιτρέπεται να θεωρηθεί αυθαίρετα «ομαδική» η σύγκρουση.
- Το επιχειρησιακό outpatient registry είναι `PATIENTS` με πραγματικό `PatientType`, ενώ recurring ραντεβού αποθηκεύονται σε `OUTPATIENT_SCHEDULE`, όχι στο απλό βοηθητικό παλιό φύλλο και όχι μόνο στο ημερήσιο κελί οθόνης. Εμφάνιση εξωτερικών με ανοιχτό μπλε περιορίζεται στους συμφωνημένους operational πίνακες `THERAPIST_DAILY`, όχι αυτομάτως σε ολόκληρη την εφαρμογή.
- `Infectious` είναι ανεξάρτητο πραγματικό operational δεδομένο. Δεν το συμπεραίνουμε από χρώμα, θάλαμο ή όνομα. Με ίσο workload/time availability μπορεί να χρησιμοποιηθεί ως tie-breaker ισόρροπης κατανομής. Ρομποτική θεραπεία ανατίθεται μόνο σε provider με αντίστοιχη πραγματική capability.
- **Φοιτητές** ξεχωριστοί από θεραπευτές. Ενεργός φοιτητής έως **5** ημερήσια patient/timeslots: στα 4/5 μπορεί να δεχτεί ακόμη ένα, στα 5/5 **όχι** έκτο. Μετά την τοποθέτηση αλλάζει η προβολή του ιστορικού με `φοιτητής 1` κ.ο.κ., διατηρώντας όμως εσωτερικά την κανονική ταυτότητα.
- **Replacement preferences:** πρώτα χαμηλότερο effective ημερήσιο φορτίο, έπειτα άμεση διαθεσιμότητα την επιθυμητή ώρα, έπειτα εναλλακτική κοινή ελεύθερη ώρα. Κατειλημμένο ζητούμενο slot δεν αποκλείει συνολικά υποψήφιο εάν υπάρχει άλλη κατάλληλη κοινή ώρα. Προτιμάται κοντινή συνεδρία ΦΘ/ανακλινόμενο, ελαχιστοποιώντας αναμονές χωρίς παράβαση hard constraints. Η 13:00 είναι χαμηλή soft preference, **όχι** απαγορευμένη ώρα.
- Ημερήσια απουσία ασθενούς απελευθερώνει το συγκεκριμένο occupied slot. Απουσία θεραπευτή ενεργοποιεί διαχείριση χωρίς απώλεια αρχικού recurring. Τα replaced sessions εμφανίζουν αρχικό ασθενή (ήπια πλάγια, **χωρίς** διαγραφή) και νέα ανάθεση/ώρα. Η διαγραφή χρησιμοποιείται για συνεδρίες που **δεν θα γίνουν**, όχι επιβεβαιωμένη αντικατάσταση.
- Πηγή λεπτομερειών: [SYSTEM_RULES.md](SYSTEM_RULES.md), [DECISIONS.md](DECISIONS.md), [DEMO_LIFECYCLE_AND_ROSTER_REQUIREMENTS_20261112.md](DEMO_LIFECYCLE_AND_ROSTER_REQUIREMENTS_20261112.md).

## 2.3 MASTER, Excel UI, Save/History και lifecycle

- Ο χρήστης δουλεύει κυρίως από το **MASTER**, με γνώριμα ελληνικά κουμπιά, φόρμες, ελεγχόμενη πλοήγηση. Τεχνικά IDs, ακατέργαστα sheets, row numbers και internal codes δεν ανήκουν σε κανονικό user-facing UI. Όπου έχουν γίνει αποδεκτά, τα gridlines, headings και tabs είναι κρυμμένα, η πλοήγηση μέσω κουμπιών· δεν πρέπει να αλλάζει αδιακρίτως όλο το VBA project.
- Ο υπάρχων `PatientProgramMasterRender.RefreshMasterPatientProgram` **λειτουργεί** για την κανονική καταχώριση και εμφανίζει την πολυγραμμική σύνθεση θεραπείας με ώρα, ημέρα, θεραπευτή στα canonical MASTER cells. **Μην ξαναγράψεις τον επιβεβαιωμένο renderer** για να διορθώσεις ξεχωριστή αρχική bulk-seed διαδρομή. Οι θεραπευτικές ειδικότητες περιλαμβάνουν ΦΘ, Ρομποτικό, Πισίνα, Ανακλινόμενο, Εργο, Λογο, ΕΦΑ, Ψυχολόγο, Απογευματινό και Κατάσταση.
- Εξιτήριο εσωτερικού και λήξη εξωτερικού είναι **date-effective γεγονότα**, όχι ημερήσια απουσία. Πρέπει να σταματούν οι μελλοντικές τοποθετήσεις, να κρατούνται οι προηγούμενες και η ακέραιη προσωπική ιστορία. `DEMO_ACTIVE_CENSUS` source path υπάρχει, αλλά το υπάρχον MASTER **δεν** κάνει ακόμη πλήρες αυτόματο live census refresh σε όλες τις in-session exit paths.
- `Κλείσιμο ημέρας` οδηγεί σε authoritative history μόνο όταν όλες οι συνεδρίες έχουν τελικό operational outcome. Η διορθωμένη κλειστή ημέρα είναι immutable version chain, παλιά version `SUPERSEDED`, νέα `CURRENT_FINAL`. Τα ιστορικά report πρέπει να μετρούν μόνο την authoritative version.
- Η πραγματική εργάσιμη μέρα θεραπευτή χρειάζεται ανεξάρτητο source. Μη μετράς «0 appointments» ως απουσία. Synthetic attendance δοκιμές H8 δεν αποτελούν κλινικά παρουσιολόγια.
- Μόνιμη αλλαγή recurring therapy απαιτεί stable `AssignmentID`, future effective-dated event, native transaction/readback και audit. Αλλαγές της ημέρας είναι ξεχωριστές.

## 2.4 Τι έχει ήδη δουλέψει πραγματικά ανά workstream

Οι παλαιότερες **operator-accepted** δοκιμές δεν ξανατρέχουν αυτόματα και τα αποδεκτά binary smokes παραμένουν FROZEN.

| Περιοχή | Πραγματική στάθμη |
|---|---|
| Core ημερήσια, απουσίες, αντικατάσταση, restore | Υπάρχουν σημαντικές source/native λειτουργίες και ιστορικά accepted paths. **Δεν** έχει γίνει ακόμη συνδυαστική πραγματική demo αποδοχή |
| H2 | `Z` native close-day/durable history **accepted/frozen** |
| H3 | `AB` native same-session correction, v1→v2, `CURRENT_FINAL` **accepted/frozen** |
| H4–H6 | Native scoped therapist/patient reports, history και contextual Save Day έχουν documented proofs |
| H7 | `BF` integrated toolbar, RefreshAll, patient/therapist reports, previous screen navigation **user accepted** για καταγεγραμμένο scope |
| H8 | `BL` patient/therapist dashboards UI **user accepted** στο συγκεκριμένο scope. `BO` 25 ημέρες **synthetic history evidence** |
| H9 | SOURCE/read-only therapist attendance/working-day contracts, **όχι** πραγματικό authoritative attendance writer ή πλήρες native chart |
| H10 | `CC` μέρος therapist charts/management visuals και native read-only/write-reopen/synthetic 179 recurring AssignmentIDs. Δεν υπάρχει πλήρης αποδοχή όλων των περιόδων/management export/permanent change event |
| Outpatient | Βασικό CREATE native green, αλλά edits/conflicts/workload έχουν επιμέρους limits. Real current outsider schedule source **pending** |
| Patients/therapists/students registry | Source validation, preview, μερικές native proofs· πλήρης ενιαία λειτουργική εγγραφή δεν έχει operator acceptance |
| Permanent recurring changes / global optimizer / 15ήμερη αξιολόγηση / broad protection | Αρκετά source/design/partial pieces. **Όχι** πλήρως εγκεκριμένο end-to-end |
| REAL-DATA ενιαίο DEMO | **ΜΟΝΟ** current code/synthetic CI και ιστορικό real-source read-only audit. **ΚΑΝΕΝΑ** real native COM build/operator acceptance |

**Αναλυτικό μητρώο/πρωτογενή τεκμήρια**: [PROJECT_DELIVERY_MASTER_REPORT_20261115.md](PROJECT_DELIVERY_MASTER_REPORT_20261115.md), [INTEGRATED_TEST_BASELINE_MATRIX_20261115.md](INTEGRATED_TEST_BASELINE_MATRIX_20261115.md), [CURRENT_HANDOFF.md](CURRENT_HANDOFF.md), `CHECKPOINT_H*`, `H10_*`. Ο παρών οδηγός κατευθύνει τη συνέχεια, οι παλιές αναφορές αποδεικνύουν το αντίστοιχο scope.

---


# ΚΕΦΑΛΑΙΟ 3. ΤΕΧΝΙΚΟΣ ΧΑΡΤΗΣ, ΑΣΦΑΛΗΣ ΜΕΘΟΔΟΣ ΚΑΙ ΑΝΤΙΜΕΤΩΠΙΣΗ ΠΡΟΒΛΗΜΑΤΩΝ

## 3.1 Υποχρεωτικός τρόπος εργασίας σε κάθε επανεκκίνηση

1. Αρχίζουμε από **αυτό το μοναδικό αρχείο**, διαβάζοντας και τα τέσσερα κεφάλαια. Ελέγχουμε μέσω GitHub connector το LIVE HEAD του `work/afternoon-medical-scope-20261115`, τα πρόσφατα commits και την τελική έκβαση των Actions. Δεν αντιμετωπίζουμε παλιό SHA ως σημερινό.
2. Συγκρίνουμε τη νέα εργασία με το Κεφάλαιο 1 (αποδεδειγμένα DONE, PENDING, BLOCKED). Δεν ξανακάνουμε την ίδια πραγματική σάρωση 106/764 χωρίς αλλαγμένη πηγή ή νέο λόγο. Τα ιστορικά H2–H10 checkpoints είναι παγωμένα τεκμήρια, όχι προγράμματα επανάληψης.
3. Εξετάζουμε ownership Python/VBA/Excel και τους κλειδωμένους κανόνες πριν αγγίξουμε κώδικα. Μικρά scoped commits, fictional-only regressions και αυστηρά αποδεικτικά για κάθε ελάττωμα. Δεν ξαναγράφουμε αποδεδειγμένα λειτουργικά MASTER/Daily/History στοιχεία εξαιτίας ελλείμματος στον χωριστό real-data initial import.
4. Τρέχουμε στο GitHub τα σχετικά tests και τα πλήρη Actions Windows + Linux μετά από αλλαγή κώδικα. Καταγράφουμε ακριβώς SHA, run ID, green/fail, αριθμό tests. Ακόμη και 2.031 πράσινα tests ΔΕΝ αποδεικνύουν native Excel ή πραγματική κλινική αποδοχή.
5. Πραγματικά patient workbooks, MRN, names, room/doctor, εμπιστευτικές διαδρομές, hashes πραγματικών patient files και κλινικά δεδομένα **ΠΟΤΕ** σε public GitHub/CI/issue/pytest fixtures. Εργαζόμαστε μόνο με aggregate non-identifying metrics δημόσια.
6. Native XLSM γράφεται αποκλειστικά με εγκεκριμένη Windows Excel/COM εκτέλεση σε νέο μοναδικό ιδιωτικό output, ποτέ με openpyxl save του πραγματικού XLSM. Original source/read-only SHA, preserved VBA, native Save/Close/Reopen, independent persisted readback και αποδεικτική operator επιβεβαίωση. Αποτυχημένο XLSM candidate παγώνει, δεν επισκευάζεται/επαναχρησιμοποιείται επιτόπου.
7. Ο assistant **προχωρά αυτόνομα** σε ασφαλή source/code/docs/CI εργασία χωρίς να περιμένει επαναλαμβανόμενα «ΟΚ». Αν απαιτείται πραγματική κλινική έγκριση, ιδιωτικό αρχείο που δεν είναι προσβάσιμο, ή Windows Excel/COM operator step, εξηγεί το *ένα* συγκεκριμένο εμπόδιο και προσφέρει προστατευμένη τοπική διαδικασία.
8. Στην ίδια εργασία ενημερώνουμε **αυτόν τον οδηγό** με νέα πράγματι verified στοιχεία και μία ακριβή επόμενη ενέργεια. Στα `CURRENT_HANDOFF`, `PROJECT_STATUS`, `ONBOARDING` αφήνουμε μικρούς δείκτες προς εδώ, χωρίς να δημιουργούμε τέσσερις ασυγχρόνιστες ενεργές οδηγίες. Οι παλιές παραδόσεις δεν διαγράφονται/αλλάζουν αναδρομικά.

## 3.2 Ακριβή paths υλοποίησης / αποδεικτικών

| Αρχείο | Χρήση |
|---|---|
| `Documentation/REAL_SOURCE_UPLOADED_PAIR_READONLY_AUDIT_20261117.md` | Πρωτογενές aggregate πραγματικό source/review proof: 106, 764, 61, **15**, F=5 U=19 V=0, 0 confirmed |
| `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md` | Υποχρεωτική πλήρης 1:1 parity και πρακτική αποδοχή |
| `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md` | Exact authorized private preflight/JSON/native CLI; παλιά blocks χωρίς `--outpatients` ΔΕΝ είναι πλήρες native build |
| `Documentation/PRACTICE_DEMO_NATIVE_SEED_RUNBOOK_20261115.md` | Native seed/check/independent readback |
| `Documentation/DECISIONS.md`, `SYSTEM_RULES.md`, `TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md` | Κλειδωμένα μοντέλα/ημερολογιακή/κλινική σημασιολογία |
| `Documentation/PROJECT_DELIVERY_MASTER_REPORT_20261115.md`, `INTEGRATED_TEST_BASELINE_MATRIX_20261115.md` | Ιστορικό workstreams, ερμηνεία βαθμίδων GREEN |
| `Documentation/IMPLEMENTATION_MAP.md`, `H10_VBA_OWNERSHIP_AND_WRITE_SURFACE_MAP_20261115.md` | Ownership και κίνδυνοι COM/VBA |
| `Python/tools/audit_real_manual_specials.py` / `rehab_excel/trial_special_fields_audit.py` | READ-ONLY F/U/V; exit 0 = κενά, 3 = special review, 2 = failed/schema; **κανένα δεν εγκρίνει import** |
| `Python/tools/audit_trial_source_alignment.py` / `rehab_excel/trial_source_alignment.py` | READ-ONLY source-row/name/cell provenance; GREEN ≠ confirmed PatientID |
| `Python/tools/audit_real_outpatient_source.py` / `trial_private_outpatient_audit.py` | Εγκεκριμένο ξεχωριστό `PATIENTS` + `OUTPATIENT_SCHEDULE`, όχι auxiliary 10 rows |
| `Python/rehab_excel/trial_identity_review.py` / `trial_admission_gate.py` | UNCONFIRMED proposals, explicit approvals, duplicate/overlap and held-source fail-closed |
| `Python/rehab_excel/trial_demo_source_reconciliation.py` / `trial_demo_source_projection.py` / `trial_import_plan.py` | Exact weekday/time/provider/provenance/read-only source hash projection |
| `Python/rehab_excel/trial_patient_registry_projection.py` / `trial_private_outpatient_projection.py` | Explicit real patient clinical identity/type and separately approved outpatients |
| `Python/tools/prepare_real_data_practice_demo.py` / `trial_demo_native_seed.py` | Full read-only preflight by default; `--build` only after all approvals, fresh eligible macro XLSM, native COM/readback |
| `Tests/test_trial_weekly_projection.py` / `test_trial_admission_gate.py` | Νέα synthetic tests `6669489` για identical-source speech duplicates, even with empty overlap sheet |

## 3.3 Αν ξαναδούμε γνωστά σφάλματα

- `can't open file ... audit_*.py`: παλιό τοπικό checkout, διορθώθηκε με fast-forward έως `3b47567` και επιβεβαιώθηκαν δύο `Test-Path=True`. Δεν είναι clinical fault.
- `else not recognized`: PowerShell split if/else στην κονσόλα· εκτέλεση ενιαίου block, όχι μεμονωμένου else.
- `MANUAL_NAME_MATCH=False` ή 0 exact-name matches: πραγματικό manual έχει **διπλό κενό μετά ΠΡΟΓΡΑΜΜΑΤΑ και κενό πριν .xlsx**. Δεν μετονομάζουμε την πηγή, χρησιμοποιούμε file-picker/verified path.
- `argparse --programme expected one argument`: κενό PowerShell `$p` μετά από επιτυχή SAFETY STOP· δεν έγινε source read. Ατομικό fail-closed block.
- `11 overlaps`: μόνο worksheet, **15 συγκρούσεις πραγματικά** στον πλήρη ανεξάρτητο έλεγχο (4 duplicate speech). Μη δημιουργήσεις κρυφή αυτόματη deduplication.
- `33 suggested PatientIDs`: προτάσεις, **0 confirmed**, όχι 33 εγκεκριμένοι.
- `F=5/U=19/V=0`: ειδικά, δεν παραλείπονται ούτε μηδενίζονται για να παρακαμφθούν οι κόφτες. Το V=0 μόνο για αυτό το snapshot.
- `10 auxiliary outpatient rows`: όχι 10 verified outpatients ούτε sessions· χρειάζεται source `PATIENTS` + `OUTPATIENT_SCHEDULE`.
- Existing MASTER canonical multi-line renderer `PatientProgramMasterRender.RefreshMasterPatientProgram` είναι ήδη λειτουργικός. Το χωριστό initial bulk seed **δεν** αποτελεί απόδειξη renderer defect.
- `CI GREEN`: όχι native XLSM, COM, κλινικό review ή πραγματική operator approval.

---

# ΚΕΦΑΛΑΙΟ 4. ΑΚΡΙΒΗΣ ΣΥΝΕΧΕΙΑ, STOP/GO, ΑΠΟΣΤΑΣΗ ΑΠΟ DEMO

## 4.1 Η αμέσως επόμενη εργασία

**Όχι ξανά το ίδιο 106/764 audit.** Το επόμενο χρήσιμο παραδοτέο είναι μία ασφαλής **ιδιωτική διαδικασία/νέο review copy**, όπου παραμένουν ρητά UNCONFIRMED μέχρι κλινική έγκριση:

1. Και οι **106** identity decisions, συμπεριλαμβανομένων των **33** proposed old IDs, με μοναδικό stable PatientID και σωστή πηγαία γραμμή.
2. Και οι **61** rows του `ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ`, οι **15** independent conflict groups (11 mixed specialties + 4 identical-source ΛΟΓΟΘΕΡΑΠΕΙΑ), οι **30** missing context rows, και κάθε source cell/year-week ambiguity.
3. Οι **5 F** ανακλινόμενες ώρες με approved 30-minute recurrence/day/provider/overlap mapping και οι **19 U** ιστορικές απογευματινές με explicit disposition. Καμία αθόρυβη παράλειψη. V=0 only for historical snapshot.
4. Η επικαιρότητα του manual της **10/11/2026** για την εβδομάδα που θα δοκιμαστεί. Αν υπάρχει νεότερο τρέχον, νέα version-bound συμφωνία χωρίς να καταστραφεί η ιστορική απόδειξη.
5. Εξωτερικοί ασθενείς από **ξεχωριστή πραγματικά εγκεκριμένη** πηγή με τα υπαρκτά canonical `PATIENTS`/`OUTPATIENT_SCHEDULE` headers, ελεγμένη read-only. Δεν κατασκευάζουμε πρόγραμμα από τις 10 auxiliary γραμμές.

**Πρακτική σειρά για τον assistant:** πρώτα εξετάζει αν υπάρχει στον κώδικα ασφαλής μη-κλινικός operator-review workflow που χρειάζεται μόνο βελτίωση (π.χ. νέα διαγνωστική αναφορά ή private new-copy checklist). Γράφει με synthetic tests και CI, χωρίς ασθενείς σε Git. Ύστερα ζητά μόνο το απολύτως απαραίτητο sign-off από εξουσιοδοτημένο χειριστή/κλινικό.

## 4.2 Μετά τις εγκρίσεις: η υποχρεωτική αλυσίδα

1. Ένα πραγματικά **current** manual snapshot + approved per-patient `review_confirmed.xlsx`, αποφάσεις για ΟΛΑ τα source holds/overlaps/F/U/V, clinical facts (type/infectious/room/doctor/MRN μόνο όταν είναι γνωστά και εγκεκριμένα).
2. Ιδιωτικό `clinical_approvals.json` σύμφωνα με το strict schema του [operator CLI](REAL_DATA_DEMO_OPERATOR_CLI_20261116.md) και επαληθευμένο **κενό αλλά πλήρες VBA** `eligible_empty_shell.xlsm`. `Rehab_Center_System_v27_1.xlsm` immutable, frozen H8/H10 proofs δεν είναι άδεια seed targets.
3. Εγκεκριμένο ξεχωριστό outpatient XLSX, strict READ-ONLY `audit_real_outpatient_source.py`. Έπειτα `Python/tools/prepare_real_data_practice_demo.py` **ΧΩΡΙΣ --build**, με `--programme --review --shell --clinical-approvals --outpatients`, same SHA files and source parity. Nonzero = STOP. Προσοχή: στο τρέχον CLI populated F/U/V προκαλούν safety stop. Οποιαδήποτε εγκεκριμένη νέα semantics/εξαίρεση απαιτεί τεκμηριωμένο, ελεγμένο **code change** με source provenance· όχι απόκρυψη πηγαίων κελιών.
4. ΜΟΝΟ τότε authorized native Windows Excel/COM `--build --output <PRIVATE_NEW_UNIQUE_DEMO>.xlsm` έξω από Git με νέο μη υπάρχον όνομα, Save/Close/Reopen, independent readback, source hashes + VBA preservation. Αν αποτύχει: freeze το συγκεκριμένο candidate και νέο όνομα μετά από διάγνωση· ποτέ overwrite.
5. Πλήρης **ΔΕ–ΠΑ** αντιπαραβολή ασθενή/ειδικότητας/ημέρας/ώρας/παρόχου με το ακριβώς ίδιο εγκεκριμένο manual, MASTER + weekly + therapist daily. Δοκιμές σε απομονωμένο αντίγραφο: νέος/εξιτήριο/λήξη, απουσία ασθενούς, απουσία θεραπευτή, replacement, θεραπευτής/φοιτητής, αποθήκευση/κλείσιμο/διόρθωση ημέρας, history/stats, rollback/reset, προστασία φύλλων, Save/Close/Reopen. Καταγραφή ανά σενάριο PASSED / FAILED / NOT TESTED με πραγματική operator απόδειξη.
6. Μόνον εφόσον παραπάνω gates πράγματι περάσουν, ονομάζουμε το πραγματικό δοκιμαστικό Excel **παραδομένο/αποδεκτό**, χωρίς να ισχυριζόμαστε ότι έγινε η τελική έκδοση εφαρμογής.

## 4.3 STOP/GO οδηγός

| Result | Ερμηνεία | Ενέργεια |
|---|---|---|
| 106/764 source matched, SHA stable | ΠΡΑΓΜΑΤΙΚΟ READ-ONLY SOURCE GREEN | Ναι στο ιδιωτικό review, ΟΧΙ στην εισαγωγή |
| 0/106 confirmed IDs | Ταυτοποίηση ΟΧΙ approved | STOP native |
| 61 holds / 15 conflicts / F 5 U 19 | Κλινικά source decisions εκκρεμούν | STOP native, ειδική per-row έγκριση |
| Auxiliary outpatients only | Όχι canonical standalone approved source | STOP complete demo |
| Preflight returns nonzero | Safety gate failed | Μη γράψεις Excel, κράτα aggregate error |
| 2.031 Windows/Linux pytest green | Source regression evidence | Δεν είναι native acceptance |
| Native COM output + independent reopen | Private **candidate** verified | Ακολουθεί πραγματικό 1:1 και UI testing |
| User full workflow/site proof | Operator-accepted DEMO scope | Μόνο τότε final scoped delivery |

## 4.4 Πόσο απέχουμε από το DEMO

**Δεν υπάρχει έντιμο μοναδικό ποσοστό** καθώς διαφορετικά gates έχουν διαφορετικό βάρος και άνθρωπο-εξαρτώμενα κλινικά approvals. Σε **οκτώ ονομαστικά milestones** έχουν πράγματι κλείσει **δύο**: (1) source/fictional code regressions και (2) read-only αντιστοίχιση των δύο παλαιών πραγματικών XLSX. Απομένουν **έξι ουσιαστικές πύλες**: (3) βεβαιωμένο current source, (4) PatientID confirmations, (5) 61 holds/15 collisions/F/U, (6) εγκεκριμένη outpatient+clinical data+clean shell, (7) native COM XLSM/save/reopen και (8) πλήρης Mon–Fri operator acceptance.

Το «2 από 8» είναι **αριθμός milestones, όχι 25% ολοκλήρωσης εργασίας**. Πολύς μηχανισμός Excel είναι ήδη μερικώς/native αποδεδειγμένος, αλλά το λειτουργικό ενιαίο **real-data DEMO παραμένει ουσιαστικά μακριά από αποδοχή**, κυρίως λόγω μη επιβεβαιωμένων πραγματικών ταυτοτήτων/θεραπειών και επειδή η πραγματική native ενοποίηση δεν εκτελέστηκε. Ακριβής ημερομηνία ή ποσοστό χωρίς κλινικές εγκρίσεις και πραγματικό Windows COM proof θα ήταν εικασία.

## 4.5 Όταν ξαναπεί ο χρήστης «βρες το αρχείο-οδηγό»

**ΜΟΝΑΔΙΚΗ ΟΡΘΗ ΑΠΑΝΤΗΣΗ ΕΝΕΡΓΕΙΑΣ:** Άνοιξε `Documentation/DIOMIL_MASTER_CONTINUITY_GUIDE.md` στο GitHub branch, διάβασε **όλα** τα κεφάλαια, επιβεβαίωσε νεότερο HEAD/CI/real evidence, εξήγησε συνοπτικά τελευταίο πραγματικό DONE και ελάχιστο πραγματικό NEXT, προχώρα στο μη-κλινικό engineering χωρίς να ρωτάς συνεχώς «οκ». **ΜΗΝ** αναζητήσεις τέσσερα άλλα ενεργά αρχεία ή αρχίσεις από την αρχή.

Στην ίδια εργασία ενημέρωσε αυτό το ένα αρχείο και πρόσθεσε μικρούς pointers, όχι τέσσερα ανεξάρτητα status narratives. **Δεν** δημοσιεύουμε πραγματικά κλινικά δεδομένα, δεν παρακάμπτουμε STOP για να «βγει demo», δεν αγγίζουμε frozen XLSM. Το παλιό `3825085` και οι 2.031 tests είναι pre-guide verified baseline και όχι υποχρεωτικά το σημερινό HEAD.

**Τελικός κανόνας:** Δεν ξαναψάχνουμε όσα έχουν αποδειχθεί, δεν εφευρίσκουμε clinical approvals, δεν ονομάζουμε demo πραγματικό πριν ολοκληρωθεί native/user acceptance.


## 4.6 ΡΗΤΗ ΠΑΥΣΗ ΟΛΩΝ ΤΩΝ PC-DEPENDENT ΔΟΚΙΜΩΝ — 17/11/2026

**ΝΕΑ ΟΔΗΓΙΑ ΧΡΗΣΤΗ:** «Κράτα αυτά τα τεστ που χρειάζονται το PC μου για το τέλος, θα μιλάμε από κινητή συσκευή μέχρι να επανέλθω στο PC. Όταν επιστρέψω, θα ενημερώσω». **ΜΗΝ ζητάς τώρα ούτε PowerShell commands, ούτε τοπικό git pull, ούτε χειρισμό native Excel, screenshots ή ιδιωτικές εγγραφές.** Επανενεργοποίησε το παρακάτω checklist αποκλειστικά όταν ο χρήστης πει ότι επέστρεψε στο PC. Αυτός ο κανόνας ισχύει για όλες τις επόμενες συνομιλίες, μέχρι νέα ρητή οδηγία.

**ΤΙ ΜΠΟΡΕΙ ΝΑ ΓΙΝΕΤΑΙ ΑΥΤΟΝΟΜΑ ΕΝΩ Ο ΧΡΗΣΤΗΣ ΕΙΝΑΙ ΣΕ ΚΙΝΗΤΟ:** GitHub connector read/write, review των πραγματικών code paths, μικρές source-only διορθώσεις, **fictional** unit/regression tests, GitHub-hosted CI Windows/Linux, μη ταυτοποιητική επεξεργασία των ήδη εγκεκριμένων aggregate ευρημάτων, ενημέρωση του ενός αυτού οδηγού με commits/αποδείξεις. Hosted Windows runner ≠ προσωπικό PC του χειριστή· test success ≠ operator COM proof. Μην προγραμματίζεις εργασία στο προσωπικό PC και μην υποσχεθείς background delivery.

**Επιβεβαιωμένο νέο cloud code checkpoint:**
- Commit [6d6b831](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/6d6b831433534c7421db60e0f1110d4a528723dc) προσθέτει **3 fictional-only regressions** σε §Tests/test_trial_admission_gate.py§: ίδιο source speech slot με αναπαραστάσεις ώρας §11:30§/§11.30§, cross-specialty ίδιος ασθενής/ώρα με κενό overlap tab, και αρνητικό control διαφορετικών ασθενών στην ίδια ώρα χωρίς ίδιο provider.
- GitHub Actions [43978249084](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43978249084) στο ακριβές SHA **4/4 jobs SUCCESS**. Full tests: **2.034 passed Windows**, **2.034 passed Ubuntu**. Τα τελευταία tests πέρασαν· **δεν** έχει πραγματοποιηθεί native Excel/clinical confirmation.
- Ο προηγούμενος οδηγός και τα πέντε γνωστά σημεία εισόδου συνδέονται με το ένα canonical αρχείο από commit §fe289b1§. Μην ξαναδημιουργήσεις τέσσερις χωριστούς οδηγούς.

**ΕΝΙΑΙΟ BACKLOG ΓΙΑ ΟΤΑΝ ΕΠΙΣΤΡΕΨΕΙ Ο ΧΡΗΣΤΗΣ ΣΤΟ PC (όχι προς εκτέλεση τώρα):**
1. Επαλήθευση/ασφαλής fast-forward του συγκεκριμένου local branch, μόνο μετά από status check· ποτέ force/reset/clean ή απώλεια τοπικών αλλαγών. Δεν είναι απαραίτητο να επαναληφθούν οι παλιοί έλεγχοι αν το PC ήδη συγχρονίστηκε.
2. Επιλογή εγκεκριμένου **τρέχοντος**, όχι αυτόματα του 10/11 historical, manual A:V και κλινικά εγκεκριμένης review copy. Διατήρηση προτάσεων ξεχωριστά από confirmed IDs, χωρίς private data σε Git.
3. Εξουσιοδοτημένη κλινική εξέταση 106 IDs, 61 holds, **15** conflict groups (11 + 4), 5 F και 19 U, ιστορικό V=0, 30 context gaps· καμία σιωπηλή απώλεια. Ιδιωτικό review copy, όχι overwriting originals. Η κλινική απόφαση **δεν** γίνεται αυτόματα από assistant.
4. Ξεχωριστή αξιόπιστη outpatient πηγή §PATIENTS§ + §OUTPATIENT_SCHEDULE§, όχι η auxiliary 10-row λίστα, πλήρες §clinical_approvals.json§ και πραγματικά καθαρό eligible Excel/VBA shell. Και τα τρία παραμένουν ανοιχτά.
5. Μόνο εφόσον όλα τα παραπάνω εγκριθούν: read-only authoritative §prepare_real_data_practice_demo.py§ preflight με §--outpatients§. Οποιοδήποτε safety stop σημαίνει **καμία** νέα native εγγραφή, όχι παράκαμψη.
6. ΜΟΝΟ σε full preflight GREEN: δημιουργία **νέου, μοναδικού και ιδιωτικού** §*_DEMO_*.xlsm§ σε Windows Excel/COM, native Save/Close/Reopen, independent reader/hash, μετά πλήρη σύγκριση ΔΕ–ΠΑ και δοκιμή των διαδρομών admissions/exit/absence/replacement/registries/history/protection/reset. Frozen baseline και παλιά πράσινα/fail artifacts αμετάβλητα.

**ΑΚΡΙΒΕΣ ΣΗΜΕΙΟ ΣΥΝΕΧΕΙΑΣ ΚΑΤΑ ΤΟ MOBILE-ONLY MODE:** Μετά τα 2.034 cloud tests, συνεχίζουμε μόνο με ασφαλείς μη-κλινικές source-code εργασίες και ανάλυση του private review workflow/15 collisions, χωρίς απαίτηση πρόσβασης σε τοπικό PC. **Το πρώτο πραγματικό Windows Excel/COM και user acceptance μεταφέρονται στο τέλος, όταν ο χρήστης επιστρέψει και το δηλώσει ο ίδιος.**


## 4.7 Νέος cloud-only κύκλος ασφάλειας review και εβδομαδιαίων PatientIDs — 17/11/2026

**Αίτημα χρήστη:** «Συνέχισε ... με όποια τέστ δεν χρειάζεται να είναι στο pc μου». Πραγματοποιήθηκαν αποκλειστικά αλλαγές κώδικα και fictional-only tests στο GitHub. Καμία τοπική εντολή, πραγματικό patient data, Windows Excel/COM, αλλαγή MASTER ή αποδοχή DEMO. Τα tests εκτελέστηκαν σε **GitHub-hosted** Windows/Linux runners, όχι στον προσωπικό υπολογιστή.

**Πρώτο commit:** [194bb419](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/194bb419ee4210c73fa9e12b7a1a4ebc535dc4ba), αρχείο `trial_admission_gate.py` και `test_trial_admission_gate.py`.
- Ξεχωριστή αυστηρή επιβεβαίωση **κάθε** παραπομπής original A:V cell στο εγκεκριμένο source row, την ειδικότητα και την υποχρεωτική source clock-column. Λανθασμένη γραμμή/στήλη, απουσία clock-column και διπλά references δεν γίνονται admissible.
- Επιλογή ακριβώς του clock cell ως source provenance ανεξάρτητα από τη σειρά στην οποία δίνονται οι έγκυρες παραπομπές. Δεν επανερμηνεύεται η ώρα ή η ειδικότητα.
- Μοναδική φυσική source row ανά ασθενή παγκοσμίως στο πρωτότυπο A:V, ανεξάρτητα από διαφορετικό clinic label. Μη ακέραιη, Boolean ή μη αριθμητική source row απορρίπτεται fail-closed.
- **4 νέες test functions** με αποκλειστικά φανταστικά ονόματα/IDs, περιλαμβάνουν αρνητικά και θετικά controls και read-only input bytes.
- GitHub [Actions 43980491249](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43980491249) **4/4 SUCCESS**, **2.038 passed Windows**, **2.038 passed Linux**.

**Δεύτερο commit:** [14d09536](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/14d095367845cbf0c60aacf84d63c07bd303e899), αρχεία `trial_weekly_projection.py` / `test_trial_weekly_projection.py`.
- Επιβεβαιωμένες ταυτότητες που διαφέρουν μόνο σε πεζά/κεφαλαία δεν αντιμετωπίζονται ως διαφορετικά PatientIDs. Μη-text/unhashable ID απορρίπτεται με ελεγχόμενο `TrialWeeklyProjectionError`, όχι ανεξέλεγκτο `TypeError`.
- **2 νέα fictional-only tests**. Τελικό GitHub [Actions 43980917796](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43980917796) στο ακριβές SHA: **2.040 passed Windows και 2.040 passed Linux, 4/4 jobs SUCCESS**. 24 pytest warnings ανά λειτουργικό, **0 test failures**. Κανένα εγκεκριμένο προηγούμενο native binary δεν μεταβλήθηκε.

**Διατήρηση ορίων:** Οι πραγματικές 106 identities **0 confirmed**, 61 source holds, 15 collision groups (11 + 4), F=5/U=19/V=0, 30 context gaps, source freshness/outpatient canonical source, private clinical approvals, clean XLSM και operator native DEMO **όλα παραμένουν εκκρεμή**. Δεν επαναλάβαμε το αληθινό 106/764 source scan και δεν δημοσιεύσαμε private identifiers ή hashes. Οι συγκεκριμένες source αλλαγές αποτρέπουν λανθασμένη **μελλοντική** admission σε ανεξάρτητες συναρτήσεις, δεν αποδεικνύουν ότι το τρέχον review είναι έγκυρο.

**Συνέχεια αποκλειστικά από κινητό:** Κρατάμε όλα τα steps με προσωπικό PC ως τελικό backlog της §4.6. Για επόμενα remote tests εξετάζουμε αν το source-only private review workflow διατηρεί τα 61/15/F/U holds αμετάβλητα και τα explicit identity decisions χωρίς προεπιλεγμένες εγκρίσεις. Οποιαδήποτε νέα πραγματική κλινική απόφαση ανήκει σε εξουσιοδοτημένο operator, όχι σε cloud tests.


## 4.8 Νέος cloud-only κύκλος διατήρησης εκκρεμοτήτων και κάλυψης ολόκληρης πηγής — 17/11/2026

**Αίτημα χρήστη:** «Πάμε», συνέχεια της §4.7 χωρίς προσωπικό PC. Δύο μικρά source-only commits, 12 επιπλέον αποκλειστικά fictional pytest cases. Δεν ανοίχθηκε/τροποποιήθηκε πραγματικό XLSX, XLSM, MASTER ή VBA. Δεν πραγματοποιήθηκε η ιδιωτική κλινική έγκριση.

**Commit [c800100](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/c800100a105728b15141ea9a4c4434d68d2878dd):** `Python/rehab_excel/trial_admission_gate.py` και `Tests/test_trial_admission_gate.py`.
- Το προηγούμενο data_only=True μπορούσε να μετατρέψει τύπο χωρίς cached value σε κενό. Ο έλεγχος τώρα διαβάζει τα raw cells και δεν αποδέχεται formula ή Excel error ως κλινικό approval, θεραπεία, shared-slot approval ή διαγραφή hold/overlap. Formula-only hold/overlap rows μετρώνται ακόμη και χωρίς εμφανιζόμενη τιμή. Το περιεχόμενο του unsafe cell δεν εμφανίζεται στα blockers.
- Οι εβδομαδιαίες προτάσεις συνδέονται μόνο με literal ρητά confirmed PatientIDs (όχι απλώς συμπληρωμένο ID), ώστε οι unconfirmed προσμετρώνται ως unadmitted. Η προεπιλεγμένη repr του admission result αποκρύπτει τα admitted entries.
- 6 νέα tests (formula-only hold, formula-only overlap, formula-generated confirmation, unconfirmed-ID weekly row, Excel error shared decision, positive result privacy). [Actions 44019556251](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44019556251): **2.046 passed Linux (28 warnings), 2.046 Windows (26 warnings), 4/4 jobs SUCCESS**.

**Commit [bd2e818](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/bd2e818a65e0286c6bd5ac0e80f374863aae2cde):** `Python/rehab_excel/trial_demo_source_reconciliation.py` και `Tests/test_trial_patient_registry_projection.py`.
- Προηγουμένως η ανεξάρτητη πηγαία ανακατασκευή εξέταζε μόνο source rows που είχαν εμφανιστεί στο review. Τώρα σαρώνεται ολόκληρη η αρχική A:V worksheet: κάθε μη κενή C:V γραμμή απαιτεί παρουσία στην εγκεκριμένη review scope. Ορφανή πληροφορία ημέρας/provider, ολόκληρη παραλειφθείσα θεραπευτική γραμμή και formula/error cells στην C:V απαιτούν ρητή ανεξάρτητη αντιμετώπιση. Δεν αγνοούνται ούτε ερμηνεύονται ως χρονικές τιμές. Γραμμή με κείμενο μόνο A:B και χωρίς C:V δεν δημιουργεί νέο ασθενή/θεραπεία αυτόματα. Αυτό είναι τεχνικό scope check για θεραπείες, όχι κλινική πιστοποίηση ότι κάθε A:B-only row αποτελεί σημείωση.
- 6 νέα pytest cases (omitted patient with therapy, orphan weekday, 3 raw formula/error variants on reviewed/unreviewed rows, nontherapy text positive control), με updated source hash ώστε να αποδεικνύεται η ίδια η business validation και όχι stale-hash rejection.
- [Actions 44019769454](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44019769454): **2.052 passed Linux (22 warnings), 2.052 passed Windows (24 warnings), 4/4 jobs SUCCESS**. Οι warnings δεν ισοδυναμούν με failures, αλλά δεν ισχυριζόμαστε ότι εξαλείφθηκαν ή ελέγχθηκαν όλες οι αιτίες τους.

**Εκκρεμότητες πραγματικών δεδομένων αμετάβλητες:** 106 ασθενείς/764 προτάσεις, 0 confirmed PatientIDs, 61 holds, 15 collision groups, F=5/U=19/V=0, 30 context gaps, freshness, outpatient source, clinical approvals, clean macro shell, native real-data Excel/COM build και operator acceptance. Δεν έγινε νέα πραγματική σάρωση. Η πολιτική raw formula/error χρειάζεται διαφανή κλινικά εγκεκριμένη αντιμετώπιση, όχι αυτόματη διόρθωση/μετατροπή στο αρχικό Excel. Δεν ισχυριζόμαστε μετρήσεις formula/error στο πραγματικό workbook.

**Επόμενο cloud-only βήμα:** ελέγχουμε τις άκρες της χωριστής outpatient source και την ενοποίηση inpatient/outpatient (ασθενής, weekday, ώρα, treatment/provider και capacity conflicts), με φανταστικά tests. Έπειτα εξετάζουμε ιδιωτικό review-copy workflow με αμετάβλητες unresolved γραμμές και μη προεπιλεγμένες εγκρίσεις. Όλο το PC backlog της §4.6 παραμένει αναβλημένο μέχρι δήλωση επιστροφής του χρήστη στο PC.


## 4.9 Κοινό πρόγραμμα εσωτερικών/εξωτερικών: έλεγχος συγκρούσεων — 17/11/2026

Συνέχεια της cloud-only εργασίας §4.8, χωρίς προσωπικό PC. Commit [7358787](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/7358787c1e2058ee8ab81f43f57f3ca1f7d2012d), τέσσερα source/test αρχεία (trial_private_outpatient_audit.py, trial_demo_source_projection.py και τα αντίστοιχα test_trial_private_outpatient_audit.py / test_trial_private_outpatient_projection.py). **9 νέα fictional tests**.

- Η χωριστή πηγή εξωτερικών ελέγχει πλέον ίδια PatientID/ημέρα/ώρα ακόμη και όταν οι ειδικότητες διαφέρουν. Διατηρείται και ο προηγούμενος duplicate-same-specialty έλεγχος. Ίδια ώρα σε διαφορετικές ημέρες δεν απορρίπτεται από αυτόν τον κανόνα.
- Η ενοποίηση των δύο πηγών ελέγχει ίδιο provider label/weekday/start-time σε inpatient-outpatient ή outpatient-outpatient εγγραφές, χωρίς σιωπηλή προσθήκη νέου μέλους σε κοινή ανάθεση. Ο ήδη εγκεκριμένος inpatient-only shared slot δεν επηρεάζεται όταν δεν διασταυρώνεται με outpatient. Τα διαφορετικά slots παραμένουν πλήρη, και το storage των εξωτερικών δεν συγχέεται με το inpatient weekly storage.
- **Όριο:** στην παρούσα real-data bulk merge διαδρομή δεν έχει υλοποιηθεί input για ρητή έγκριση outpatient/shared membership. Μια τέτοια σύγκρουση σταματά το candidate και παραπέμπεται για ανεξάρτητη έγκριση και ασφαλή υποστήριξη, δεν αποτελεί νέο καθολικό ban στις νόμιμες διπλοκρατήσεις. Δεν αλλάχτηκε ο υφιστάμενος operational double-booking μηχανισμός. Provider σύγκριση με trim/casefold των labels της πηγής μόνο· όχι τεκμήριο επίλυσης aliases, διαφοροποιημένων IDs, διάρκειας ή πλήρους capacity/availability. Δεν εικάζεται θεραπευτής όταν λείπει. Οι έλεγχοι εδώ αφορούν ίδιο start-time, όχι ολοκληρωμένη ανάλυση διαστημάτων.
- [Actions 44020691900](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44020691900): **2.061 passed Linux (27 warnings), 2.061 passed Windows (24 warnings), 4/4 jobs SUCCESS**. Cloud runners, όχι τοπικό Excel/COM. Κανένα πραγματικό αρχείο ή πληροφορία ασθενή δεν χρησιμοποιήθηκε/τροποποιήθηκε σε αυτόν τον κύκλο.

**Επόμενο cloud-only βήμα:** έλεγχος του private review-copy workflow: όλες οι εκκρεμότητες και οι προτάσεις διατηρούνται σε ξεχωριστό αντίγραφο χωρίς αυτόματη επιβεβαίωση ή απώλεια εγγραφών. Διερεύνηση ρητών shared-slot approvals για την κοινή πηγή και canonical provider identity/capacity με βάση τις ήδη εγκεκριμένες συμβάσεις, χωρίς εκ των προτέρων έγκριση. F/U/V, 106 ταυτότητες (0 confirmed), 61 holds, 15 conflict groups, 30 context gaps, canonical outpatient source, clinical approvals, clean native shell και operator acceptance παραμένουν εκκρεμή. PC backlog §4.6 αναβλημένο.


## 4.10 Ιδιωτική συγκεντρωτική λίστα εκκρεμοτήτων — 17/11/2026

Συνέχεια mobile-only §4.9, commit [26f3f18](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/26f3f1874ce47cf56c7280ed3223086ea68dbe06). Νέα αρχεία `Python/rehab_excel/trial_review_checklist.py`, `Python/tools/prepare_trial_review_checklist.py`, `Tests/test_trial_review_checklist.py`. 8 νέα fictional-only tests.

- Χρησιμοποιούμε το υπάρχον review (και τις ήδη παραχθείσες προτάσεις G:H), δεν επαναλαμβάνουμε identity matching. Επιλέχθηκε **νέο ανεξάρτητο companion XLSX**, όχι επανεγγραφή ή διακόσμηση του review, ώστε τα δύο input workbooks να διατηρούνται byte-for-byte. Μόνο σταθερές κατηγορίες, ονόματα φύλλων, διευθύνσεις/αριθμοί γραμμών, κωδικοί λόγου, κατάσταση ΕΚΚΡΕΜΕΙ και κενές σημειώσεις χειριστή. Δεν αντιγράφονται ονόματα, PatientIDs, ώρες ή αρχικές κλινικές τιμές. Η ιδιωτική καρτέλα ΟΔΗΓΙΕΣ δεσμεύει τα input SHA256 τοπικά, ποτέ στο Git ή stdout. Οι διευθύνσεις/metadata παραμένουν ιδιωτικά.
- Αναγνωρίζονται unconfirmed ή επαναλαμβανόμενες confirmed identities, ελλιπές context, source-link discrepancies, όλα τα υπάρχοντα holds/overlap rows (ακόμη και formula-only), ανεξάρτητες ομάδες ίδιας πηγής/ημέρας/κανονικοποιημένης ώρας με ίδια ή διαφορετική ειδικότητα, non-proposed/invalid weekly rows, περιεχόμενο F/U/V, raw formula/error και θεραπευτικές γραμμές που λείπουν από review. Οι ήδη καταγεγραμμένες και οι ανεξάρτητες συγκρούσεις είναι χωριστές κατηγορίες, όχι αθροιστικά μοναδικά περιστατικά.
- Ασφάλεια: distinct XLSX paths εκτός Git, exclusive creation, προστασία υφιστάμενου/ανταγωνιστικού output, check SHA256 πριν/μετά και ακριβές readback όλων των γραμμών. Σε failure δεν υπάρχει εγκεκριμένο αποτέλεσμα. CLI εκτυπώνει μόνο κατηγορία/πλήθος.
- **Όριο:** Η λίστα είναι διαγνωστική, όχι πλήρες admission preflight, not clinical approval, όχι αλλαγή σε review ID/hold/slot ακόμη κι αν ο χειριστής γράψει σημείωση. Δεν επιλύει F/U/V, shared-provider approvals, capacity, durations ή πραγματικά source data και δεν ανοίγει κανένα native gate. Δεν εκτελέστηκε στα δύο πραγματικά αρχεία: οι προηγούμενες 106/61/15/30/5/19 μετρήσεις δεν είναι νέο αποτέλεσμα checklist.
- Όταν επιστρέψει ο χρήστης στο PC, με εγκεκριμένα private paths, μόνο τότε `python Python/tools/prepare_trial_review_checklist.py --programme "<private original.xlsx>" --review "<private review.xlsx>" --output "<new private checklist.xlsx>"`. Κανένα private input/output στο Git checkout. Μην ζητήσεις την εντολή τώρα. Η λίστα αφορά συγκεκριμένο snapshot: αν αλλάξουν δεδομένα, απαιτείται νέα έκδοση και νέα λίστα, όχι overwrite.

**CI:** [Actions 44044323627](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44044323627), exact code SHA: **2.069 passed Linux (23 warnings)**, 2069 passed, 27 warnings in 21.80s. 4/4 jobs SUCCESS. Δεν χρησιμοποιήθηκε πραγματικό αρχείο ασθενών ή native Excel/COM. 

**Επόμενο cloud-only βήμα:** μελέτη explicit shared-slot approvals για το κοινό inlet και των canonical provider/capacity contracts. Να μη διαγραφούν οι ήδη εγκεκριμένες εξαιρέσεις και να μην επινοηθούν aliases, διαθέσιμες ώρες ή θεραπευτικές διάρκειες. Όλοι οι πραγματικοί κλινικοί και native PC φραγμοί §4.6 παραμένουν σε ισχύ.


## 4.11 Κοινός κανόνας ταυτότητας θεραπευτή και δυναμικότητας — 17/11/2026

Mobile-only συνέχεια §4.10. Commit [1b34311](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/1b34311b5a4331cb7bb773031848f15d5f12553d), τέσσερα τροποποιημένα source/test αρχεία και νέο `Python/rehab_excel/trial_provider_capacity.py`. 9 νέα fictional-only pytest cases.

- Επαναχρησιμοποίηση του ήδη εγκατεστημένου `rehab_core.recurring_conflicts._provider_key` τόσο στο inpatient trial admission/explicit shared-slot approvals όσο και στο mixed source merge. Ελληνικοί τόνοι, πεζά/κεφαλαία, τελείες, κενά, τελικό σίγμα σύμφωνα με τον υπάρχοντα κώδικα. Το πραγματικό label παραμένει αμετάβλητο στα source rows. Ίδιο κανονικοποιημένο κλειδί απαιτεί έγκριση ακριβούς membership στην inpatient admission· διπλή έγκριση με διαφορετική γραφή απορρίπτεται. Mixed/outpatient collision εξακολουθεί να απαιτεί ανεξάρτητη explicit approval και δεν περνά σιωπηλά. Δεν αποτελεί επίλυση aliases ή αντιστοίχιση canonical registry IDs.
- Νέα καθαρή read-only διαγνωστική `audit_trial_provider_capacity(projection, approved_limits=...)`: δέχεται ρητά εγκεκριμένα διαθέσιμα **πλήθη slots ανά provider label και weekday**, όχι legacy 6/5 fallback. Επαναχρησιμοποιεί `validate_recurring_timeslot_capacity`, μετράει διαφορετικές ώρες/ημέρα και όχι ασθενείς (δύο εγκεκριμένοι ασθενείς στο ίδιο ρολόι = μία κατειλημμένη ώρα). Ξεχωρίζει missing limits, υπέρβαση, ρητά μηδενικό όριο, και μη ανατεθειμένες θεραπείες. Απορρίπτει Boolean/αρνητικά/μη επιτρεπτή ημέρα/διπλά canonical limits. Μόνο συνολικά μεγέθη, χωρίς κλινικά values. Δεν αποδίδει υποχρεωτικά provider σε optional unassigned therapy.
- **Σημαντικό όριο:** το νέο capacity audit είναι ανεξάρτητο Python diagnostic, όχι ακόμη ενσωματωμένο σε native build ή CLI preflight, και ο caller πρέπει να παρέχει εγκεκριμένα πραγματικά όρια ανά θεραπευτή/ημέρα. Δεν είναι απόδειξη πραγματικής διαθεσιμότητας σε κάθε ακριβή ώρα, διάρκειας θεραπείας, attendance/leadership/effective overrides, clinician roster ID mapping, ή εξουσιοδότηση υπέρβασης. Δεν αλλάζουν οι υφιστάμενοι operational capacity/exception rules. Δεν υλοποιήθηκε ακόμη διαδρομή ρητής outpatient/mixed shared-slot approval στο bulk workflow.
- [Actions 44067635285](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44067635285), exact code SHA: **2.078 passed Linux (28 warnings)**. 2078 passed, 29 warnings in 28.29s 4/4 jobs SUCCESS. Κανένα πραγματικό XLSX/XLSM, VBA ή MASTER δεν τροποποιήθηκε.

**Επόμενο ουσιαστικό cloud-only βήμα:** ενιαίος έλεγχος εισαγωγής με εγκεκριμένο snapshot provider capacities/available clock slots και exact mixed-group decisions, με αυστηρή δέσμευση στα inputs και χωρίς νέο clinical default. Εξετάζουμε υπάρχουσες authoritative sources πριν προσθέσουμε νέο schema. Το §4.6 PC backlog παραμένει· 106 IDs, 61 holds, 15 collisions, F/U/V, πραγματική outpatient πηγή, κλινική έγκριση, clean native shell και operator acceptance εκκρεμούν. Μην μετατρέπεις green fictional test counts σε DEMO completion percentage.


## 4.12 Ενσωμάτωση διαθέσιμων ωρών και κοινών εγκρίσεων στο preflight/DEMO — 17/11/2026

Ρητή επιβεβαίωση χρήστη: δύο ασθενείς στον ίδιο θεραπευτή/ώρα με συμπληρωματικές ημέρες Δε-Τε-Πα και Τρ-Πε καταλαμβάνουν **μία θέση στο εβδομαδιαίο πρόγραμμα**, όχι δύο. Παραμένουν διαφορετικοί ασθενείς στο μητρώο, χωρίς ταυτόχρονη θεραπεία και χωρίς shared exception. Δύο ασθενείς την ίδια ημέρα/ώρα είναι ξεχωριστή περίπτωση και χρειάζονται ρητή έγκριση. Η λογική του υφιστάμενου operational scheduler δεν αντικαθίσταται. Mobile-only.

- Commit [42861e3](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/42861e35f3a74abc204ab90a28bea8c74810d1dc), 10 νέα fictional pytest cases. Επεκτάθηκε η source-only `merge_trial_demo_outpatients`: ρητή exact-membership έγκριση μόνο για πραγματική κοινή ομάδα με συμμετοχή εξωτερικού ασθενούς. Stale, ελλιπής (π.χ. παλιό ζεύγος + νέο μέλος) ή περιττή έγκριση απορρίπτεται. Ήδη ελεγμένες inpatient-only ομάδες παραμένουν ως είχαν. Χρησιμοποιείται η υπάρχουσα κανονικοποίηση provider label.
- Νέο `trial_provider_schedule.py`: ιδιωτικό snapshot ρητά εγκεκριμένων **ακριβών ωρών ανά θεραπευτή/ημέρα**, και explicit mixed/outpatient shared membership. Επαναχρησιμοποιεί `audit_trial_provider_capacity` και τον operational distinct-time κανόνα. Δεν θεωρεί δύο complementary weekdays διπλοκράτηση. Ακόμη και αν χωρά αριθμητικά μια νέα ώρα, εκτός εγκεκριμένου clock set δεν περνά. Missing availability για named provider απορρίπτεται. Οι optional unassigned therapies παραμένουν τέτοιες.
- Συνδέθηκε το ίδιο gate στην `prepare_real_data_practice_demo.py` και στην `seed_trial_demo_workbook` πριν από οποιαδήποτε native write. Νέα παράμετρος `--provider-schedule` για το ιδιωτικό εγκεκριμένο snapshot. Τεχνική επιλογή: χωριστό JSON, όχι αλλαγή καθιερωμένων XLSM, κλινικών sheets ή operational UI· δεν υπάρχει επιβεβαιωμένη πηγή ακριβών availability clocks για το συγκεκριμένο πραγματικό DEMO. Ίδια source hashes programme/review/shell/outpatients δένουν την έγκριση με τη συγκεκριμένη έκδοση, ελέγχεται το JSON hash πριν/μετά, και επαναλαμβάνεται gate στον native seeder. Αν αλλάξει approval JSON μέσα στο native build, δεν δηλώνεται αποδεκτό αποτέλεσμα. Χωρίς provider snapshot το read-only CLI δίνει **PARTIAL READ-ONLY SOURCE CHECK**, όχι πλήρες GREEN· native build σταματά. Προσοχή: exit 0 σε partial diagnostics σημαίνει επιτυχία ελέγχου μόνο, όχι πλήρη άδεια.
- **Όρια:** Μόνο source-bound bulk admission approval. Δεν γράφει νέο operational exception event, δεν αλλάζει/επαληθεύει το υπάρχον VBA double-booking, την επίσημη operational capacity config, τις effective-day overrides, τις διάρκειες ή την πραγματική provider ID αντιστοίχιση. Ο αριθμός εγκεκριμένων ωρών πρέπει να προκύψει από εξουσιοδοτημένο πραγματικό πρόγραμμα, όχι από τα δεδομένα ασθενών ούτε από την επιθυμία να χωρέσουν. Δεν ελέγχθηκαν τα πραγματικά αρχεία και δεν δημιουργήθηκε πραγματικό availability JSON ή DEMO. Όλα τα πραγματικά F/U/V, identity/holds, current source, outpatient source και clinical/PC acceptance gates εξακολουθούν να ισχύουν. 
- [Actions 44070526343](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44070526343): **2.088 passed Linux (25 warnings)**. 2088 passed, 29 warnings in 19.02s. 4/4 jobs SUCCESS.

**Operator:** Η αναθεωρημένη μορφή τεχνικού input/εντολής είναι στην αρχή του `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Δεν ζητείται χρήση του PC ή συμπλήρωση από κινητό. Επόμενο remote βήμα: περιορισμένη εξέταση των πραγματικών ανοιχτών φραγμών, ιδίως F (30 λεπτά, PT-linked) και V/U scope με επαναχρησιμοποίηση των υπαρχουσών συμβάσεων. Προτεραιότητα πρακτικό DEMO και όχι νέες παράλληλες υποδομές/αέναοι έλεγχοι. 


## 4.13 Ιστορικό απογευματινό U: ρητή διατήρηση εκτός εισαγωγής — 17/11/2026

Commit [40c5811](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/40c5811eff399b88d6f2c4748783f8d07514946d), 7 νέα fictional tests. Εφαρμόστηκε η προϋπάρχουσα οδηγία: U παραμένει ιστορικό περιεχόμενο στο αμετάβλητο πρωτότυπο, **ποτέ** μεταφορά σε πρωινή ΦΘ ή νέο PM slot. Στο ιδιωτικό review προστίθεται χωριστό φύλλο `ΕΓΚΡΙΣΕΙΣ_ΠΕΔΙΟΥ_U` με ακριβή headers `SourceCell, ProgrammeSHA256, Decision` και literal `ACKNOWLEDGED_NOT_IMPORTED` για κάθε πραγματικά συμπληρωμένο U. Απόφαση εξουσιοδοτημένου χειριστή, όχι auto-fill. Δεσμεύεται στο ακριβές SHA256 του ιδιωτικού πρωτοτύπου και πλήρες σύνολο διευθύνσεων, απορρίπτονται missing/extra/stale/duplicate/unsafe/formula/orphan. Τα U παραμένουν στην πηγή, όχι στη νέα εβδομαδιαία προβολή. Δεν αποτελεί αυτόματη διαγραφή παλιών hold rows ή έγκριση ταυτοτήτων. Η CLI και η ανεξάρτητη source reconciliation ελέγχουν τη διάθεση U. Η προηγούμενη καθολική απαγόρευση U αίρεται ΜΟΝΟ όταν ολοκληρωθεί αυτή η ρητή αναγνώριση. F/V παραμένουν blockers. Δεν εγκαταστάθηκε/ελέγχθηκε native PM UI ή κενό PM sheet στο πραγματικό DEMO. Το υπάρχον separate afternoon contract παραμένει.

[Actions 44086028522](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44086028522) exact code SHA: **2.095 passed Linux (27 warnings), 2.095 passed Windows (28 warnings), 4/4 jobs SUCCESS**. Δεν χρησιμοποιήθηκε πραγματικό XLSX/XLSM, ούτε δημιουργήθηκε πραγματική έγκριση για τα 19 U. Δεν τροποποιήθηκαν MASTER, VBA και operational scheduler.

**Επόμενο cloud-only βήμα:** ολοκλήρωση explicit F/V source mapping, με το υπάρχον 30λεπτο Ανακλινόμενο και morning pool/operating-hours contracts, χωρίς εικασία ημέρας, θεραπευτή, διάρκειας άλλης θεραπείας ή χωρητικότητας. Native έλεγχος μη κληρονόμησης παλιών PM records σε καθαρό shell και νέων blank PM fields. Προτεραιότητα άρση πραγματικών blockers, όχι πρόσθετες παράλληλες υποδομές. PC backlog §4.6 αμετάβλητο. Ρητό αίτημα χρήστη: ο κεντρικός οδηγός και οι αντίστοιχες πρακτικές οδηγίες να ενημερώνονται σε κάθε ολοκληρωμένο στάδιο, με επαληθευμένα αποτελέσματα, εκκρεμότητες και ακριβές επόμενο βήμα, ώστε να συνεχίζει νέα συνομιλία χωρίς επανάληψη.


## 4.14 Εγκεκριμένη εισαγωγή Ανακλινόμενου F / Πισίνας V — 17/11/2026

Commit [1aa9a06](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/1aa9a066e948315a81cf77aec0fa365b6f27c86f), 8 επιπλέον fictional pytest cases. Οριστική υπενθύμιση: δεν αλλάζουμε υφιστάμενο MASTER/VBA ή solver. Το `trial_demo_planner_mirror.py` πλέον περιλαμβάνει τις ήδη καθορισμένες στήλες `Ανακλινόμενο_Ώρα / Ανακλινόμενο_Ημέρες` και `Πισίνα_Ώρα / Πισίνα_Ημέρες`, χωρίς provider column. Ο υφιστάμενος native writer και renderer/independent readback τις καταναλώνουν. Η `trial_source_alignment.py` πλέον γνωρίζει την ΠΙΣΙΝΑ με αποκλειστικό source V. Η απόδοση των απλών εγγραφών έχει ώρα πάνω και ημέρες κάτω. Δεν εγκαταστάθηκε νέο UI.

- Ιδιωτικό review, χωριστό φύλλο `ΕΓΚΡΙΣΕΙΣ_ΠΕΔΙΩΝ_FV` με ακριβή headers `SourceCell, ProgrammeSHA256, Weekdays, Decision`. Για κάθε και μόνο συμπληρωμένο F και V ακριβής διεύθυνση, SHA256 ολόκληρου αμετάβλητου programme, εγκεκριμένες canonical ημέρες χωρισμένες με παύλα (π.χ. ΔΕ-ΤΕ-ΠΑ), literal `APPROVED_MORNING_SERVICE_SLOT`. Η απόφαση δηλώνει ότι εξουσιοδοτημένος χειριστής επιβεβαίωσε πραγματικές ώρες λειτουργίας, χωρητικότητα υπηρεσίας, πλήρη συμβατότητα intervals και τυχόν κλινική σύνδεση με ΦΘ. Δεν γίνεται auto-approval, εκχώρηση ονόματος PT, ούτε εφαρμογή implicit weekdays στα F/V. Η ώρα προέρχεται αποκλειστικά από το αρχικό κελί. Formula/error/orphan/stale/duplicate/surplus/missing approvals και ΣΚ απορρίπτονται. Η αυστηρή ισότητα του ανασυσταθέντος προγράμματος με το ανεξάρτητα εγκεκριμένο weekly review συνεχίζει να εφαρμόζεται.
- `verify_morning_special_scope` + `preview_source_weekdays` + source reconciliation + CLI gate. F διατηρεί τη γνωστή διάρκεια 30 λεπτών ως κανόνα προελέγχου: ελέγχεται ότι δεν υπερβαίνει το τέλος ημέρας και ότι δεν ξεκινά άλλη πρωινή θεραπεία του ίδιου ασθενούς μέσα στο 30λεπτο. Δεν υποτίθεται διάρκεια άλλων ειδικοτήτων ή της πισίνας. Η σαφής έγκριση της υπηρεσίας είναι αναγκαία για πραγματικά πλήρη intervals, operating hours, resource capacity και ΦΘ adjacency. Δεν δημιουργήθηκε νέος γενικός duration/capacity solver, ούτε ενσωματώθηκε εδώ η operational service-hours configuration. Το U παραμένει εκτός αυτόματης εισαγωγής (§4.13).
- **Εκκρεμότητες/όρια**: οι 5 πραγματικές F εγγραφές και 17 προτάσεις δεν εγκρίθηκαν, το πραγματικό V ήταν κενό στο ιστορικό snapshot. Καμία πραγματική κλινική ή resource έγκριση/μεταβολή, κανένα XLSM. Native readback/visual/functional acceptance και resource constraints ακόμη απαιτούνται. Το γνωστό 30min δεν γράφεται ως νέο duration field στο υπάρχον weekly schema. Προηγούμενα hold/overlap rows δεν αποσύρονται αυτόματα. Standalone special audit παραμένει inventory, όχι αποτέλεσμα admission.

**CI** [Actions 44101704459](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44101704459): 2.103 passed Linux (25 warnings); 2103 passed, 25 warnings in 28.61s. 4/4 jobs SUCCESS.

**Επόμενο συγκεκριμένο cloud-only βήμα:** αποτροπή κληρονόμησης παλιών `AFTERNOON_PHYSIO_SLOTS` ή λοιπών afternoon state σε clean shell και εγγύηση κενής ανεξάρτητης PM πηγής, με επαναχρησιμοποίηση του υπάρχοντος schema / installer contract. Κατόπιν συγκέντρωση ελάχιστου ιδιωτικού operator checklist για πραγματικά approvals, όχι νέα επαναλαμβανόμενα diagnostics ούτε synthetic DEMO. PC backlog §4.6 παραμένει έως ρητή επιστροφή του χρήστη.


## 4.15 Κενή ανεξάρτητη απογευματινή πηγή στο DEMO — 17/11/2026

Commit [94db7c4](https://github.com/diomil73/EKA-REHAB-TIMETABLE/commit/94db7c48640c6701dabf9a08bc00e83c735e857c), 10 νέα συνθετικά tests. Ο native seeder επαναχρησιμοποιεί το υπάρχον `AFTERNOON_PHYSIO_SLOTS` (PatientID, Treatment, Weekday, StartTime, Therapist). Καθαρό shell: επιτρέπεται απουσία φύλλου ή ακριβές header-only veryHidden φύλλο. Παλιό περιεχόμενο σε οποιαδήποτε θέση, raw formula/error, αλλοιωμένα headers, πρόσθετη επικεφαλίδα ή λάθος ορατότητα προκαλούν STOP, όχι καθαρισμό. Μόνο στο νέο staged DEMO δημιουργείται το κενό φύλλο αν λείπει. Ανεξάρτητη raw readback απαιτεί το κενό φύλλο, σωστά headers και veryHidden. Ούτε U ούτε πρωινή ΦΘ αντιγράφονται στο PM. Δεν άλλαξαν MASTER, VBA ή scheduler.
[Actions 44102691014](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44102691014): **2.113 passed Linux (27 warnings), 2.113 passed Windows (28 warnings), 4/4 jobs SUCCESS**. Native πραγματικό Excel/COM δεν εκτελέστηκε. Το PM storage initialization δεν εγκαθιστά writable PM UI, δεν ολοκληρώνει την persistence/medical deferral/replacement/history λειτουργικότητα του PM. Η υπάρχουσα read-only ένδειξη και validated pure contracts δεν ισοδυναμούν με end-to-end PM.


## 4.16 ΕΝΙΑΙΑ ΤΕΛΙΚΗ ΠΡΟΕΤΟΙΜΑΣΙΑ ΠΡΑΓΜΑΤΙΚΟΥ DEMO — 17/11/2026

**CURRENT RESUME / GO-NO-GO.** Ο χρήστης ζήτησε να ολοκληρωθεί τώρα μόνο το πρώτο στάδιο, η προετοιμασία. Ο παρών οδηγός αποτελεί το ένα σημείο συνέχειας. Η cloud προετοιμασία των οδηγιών ολοκληρώθηκε· ΟΧΙ οι πραγματικές εγκρίσεις, η κατασκευή ή η αποδοχή. **STOP για PC actions μέχρι ρητή επιστροφή του χρήστη (§4.6).** Δεν χρειάζεται νέος ατέρμονος κύκλος συνθετικών tests ή νέα παράλληλη αναφορά. Ακριβής επόμενη ενέργεια τότε: επιλογή εγκεκριμένων πραγματικών πηγών και ασφαλής συγχρονισμός, όχι αυτόματα build. Οι προηγούμενες ενότητες με blanket F/U/V stop έχουν αντικατασταθεί από §§4.13–4.15 για τις ρητά εγκεκριμένες περιπτώσεις. Το 2.113/2.113 είναι code regression evidence, όχι ποσοστό ολοκλήρωσης.

### Α. Έξι ιδιωτικά inputs (όλα διαφορετικά, εκτός Git)

| Input | Απαραίτητη προϋπόθεση | Κατάσταση τώρα |
| --- | --- | --- |
| 1. programme XLSX | Πραγματικό, τρέχον και εγκεκριμένο A:V, αμετάβλητο πρωτότυπο | Ιστορικό 10/11 ελέγχθηκε read-only, όχι επιβεβαίωση τρέχοντος |
| 2. review XLSX | Νέα ιδιωτική εγκεκριμένη αναθεώρηση του αντίστοιχου review, ίδιο source snapshot, πλήρως επιβεβαιωμένες ταυτότητες και εβδομαδιαίες εγγραφές | Εκκρεμεί |
| 3. clinical approvals JSON | Ακριβής clinic/source_row, πραγματική ταξινόμηση και λοιπά κλινικά δεδομένα | Εκκρεμεί |
| 4. outpatients XLSX | Χωριστή εγκεκριμένη canonical PATIENTS + OUTPATIENT_SCHEDULE | Εκκρεμεί· ο βοηθητικός 10-row πίνακας δεν αρκεί |
| 5. eligible empty shell XLSM | Πραγματικό συμβατό Excel/VBA, μηδέν κλινικά/legacy history/PM δεδομένα· πλήρης βασική εφαρμογή | Εκκρεμεί· όχι το παλιό populated v27_1 ή frozen H2–H10 |
| 6. provider schedule JSON | Εγκεκριμένες ακριβείς διαθέσιμες ώρες, πραγματικά όρια και ακριβείς κοινές αναθέσεις, δεμένο με τελικές εκδόσεις programme/review/shell/outpatients | Εκκρεμεί, δημιουργείται τελευταίο μετά το κλείδωμα των 4 inputs |

Τοπικά, εξουσιοδοτημένα αρχεία, ποτέ Git/CI/chat με patient rows, names, IDs, clinic values, ιδιωτικές διαδρομές ή source hashes. Χρησιμοποιούμε selection πλήρους πραγματικής διαδρομής, όχι brittle filename search. Το γνωστό ιστορικό αρχείο έχει διπλό κενό μετά το ΠΡΟΓΡΑΜΜΑΤΑ και κενό πριν από .xlsx (§1.3). Δεν χρειάζεται μετονομασία ούτε νέα αναζήτηση του ήδη γνωστού ζεύγους, αλλά επιβεβαίωση τρέχουσας πηγής.

### Β. Μία ουσιαστική ιδιωτική διαδικασία έγκρισης

1. Διατήρησε αρχικό programme και αρχικό review suggestions αμετάβλητα. Φτιάξε μία νέα αναθεώρηση review σε ιδιωτική θέση. Εφόσον οι πηγές είναι διαθέσιμες και έχει επιλεγεί η πραγματική έκδοση, ο ήδη διαθέσιμος `prepare_trial_review_checklist.py --programme ... --review ... --output ...` δημιουργεί νέα ιδιωτική companion checklist (όχι approval). Μην ξανατρέχεις για ίδιο snapshot αν υπάρχει ήδη ελεγμένο checklist. Αν αλλάξει snapshot, ανανέωσε το review/αποδείξεις. Το ιστορικό baseline: 106 review patients, 764 προτεινόμενες placements, 0 confirmed IDs, 33 old-ID suggestions, 61 holds, 15 same-patient collision groups (11 + 4 πρόσθετα), 30 context gaps, 5 F / 17 προτάσεις, 19 U, V=0. Οι κατηγορίες ενδέχεται να επικαλύπτονται· όχι άθροισμα μοναδικών προβλημάτων. Αυτοί ΔΕΝ είναι στόχοι που πρέπει να επιβληθούν στο νέο τρέχον snapshot.
2. Επιβεβαίωσε κάθε πραγματικό PatientID στο φύλλο ΑΣΘΕΝΕΙΣ (στήλες E/F: explicit ID + `CONFIRMED`). Πραγματικά μοναδικό ID, ίδιο source row/clinic/name/context. Οι προτάσεις G/H δεν είναι επιβεβαίωση. Μην συμπληρώνεις A κενό από προηγούμενη γραμμή. Όλα τα πραγματικά θεραπευτικά source rows πρέπει να καλύπτονται. Οι εγγραφές της inpatient πηγής οφείλουν να ανήκουν στο inpatient πρόγραμμα· αν η κλινική επιβεβαίωση δείχνει διαφορετικό scope, απαιτείται ορθή εγκεκριμένη οργάνωση της πηγής, όχι ψευδής ταξινόμηση.
3. Επανεξέτασε κάθε hold/overlap με ακριβή αρχική τιμή και κλινική απόφαση. Ο ίδιος ασθενής δεν κάνει δύο θεραπείες ταυτόχρονα. Δεν απορρίπτουμε έγκυρη θεραπεία για να περάσει το test. Καμία σιωπηλή διαγραφή ή dedup. Αν η ίδια η αρχική πηγή περιέχει πραγματική ασυνέπεια, ζητείται νέα εξουσιοδοτημένη διορθωμένη έκδοση ως νέα πηγή, διατηρώντας το πρωτότυπο και νέο hash· δεν επιτρέπεται απλή αλλαγή ώρας στο review. Κράτα σε ξεχωριστό ιδιωτικό αρχείο τις αποφάσεις και το πλήρες ιστορικό των holds/overlaps· τα ενεργά sheets ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ και ΕΠΙΚΑΛΥΨΕΙΣ στη νέα approved revision must have no populated data rows **μόνο αφού** έχουν πραγματικά επιλυθεί. Η gate μετρά οποιαδήποτε ενεργή data row, όχι ένα status flag. Η weekly ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ διατηρεί ακριβή original ώρα/day/provider/reference, μία placement ανά weekday, literal `PROPOSED_EXACT_SOURCE` (τεχνικός marker, όχι μόνος του clinical approval). Blank clock=no booking; blank day=ΔΕ–ΠΑ μόνο για τα συνηθισμένα A:V time/day fields. Καμία θεραπεία ΣΚ. Ο provider στις G/H weekly στήλες: απαιτούμενος για ΦΘ/Ρομποτικό στη G, optional name στη H για τις άλλες ειδικότητες εφόσον υπάρχει στο source. F/V χωρίς επινοημένο όνομα.
4. Ειδικές θεραπείες: F/V approved days και πραγματική υπηρεσία/interval/capacity από εξουσιοδοτημένο άνθρωπο. Ίδια source ώρα, F 30 λεπτά και ΦΘ-linked, V πρωινό, καμία εικασία ημερών. Φύλλο ΕΓΚΡΙΣΕΙΣ_ΠΕΔΙΩΝ_FV: SourceCell / ProgrammeSHA256 / Weekdays / Decision, μία γραμμή ανά populated F/V, ακριβές literal `APPROVED_MORNING_SERVICE_SLOT`. Επαρκές scope δεν αποτελεί αυτόματη επίλυση πλήρους interval/resource solver. U: ΕΓΚΡΙΣΕΙΣ_ΠΕΔΙΟΥ_U, SourceCell / ProgrammeSHA256 / Decision, literal `ACKNOWLEDGED_NOT_IMPORTED` για κάθε U, ποτέ αυτόματη νέα PM θεραπεία. Βλέπε §§4.13–4.14. Κανένα άδειασμα αρχικών U/F/V.
5. Ακριβής κοινή ανάθεση εσωτερικών μόνο: ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS με Therapist / Weekday / StartTime / PatientIDs / Decision, exact membership IDs χωρισμένα με |, literal `APPROVED_SHARED`. Mixed/outpatient groups στο provider JSON, όχι duplicate approval διαδρομή. Καμία υπόθεση γενικής άδειας διπλοκράτησης. Δε-Τε-Πα και Τρ-Πε στην ίδια ώρα είναι **ένα weekly timeslot**, όχι ταυτόχρονη κράτηση. Ξεχωριστή ίδιας ημέρας distinct-time capacity και ακριβές availability clock gate. Επιβεβαίωσε effective roster identities, actual διαθέσιμες ώρες, approved exceptions. No invented aliases/availability.
6. Clinical JSON schema `diomil-private-clinical-approvals-v1`, root `schema,records`. Ένα record ανά ακριβές clinic/source_row με patient_type (INPATIENT ή OUTPATIENT, πραγματική εγκεκριμένη ταξινόμηση), infectious (JSON true/false μόνο με κλινική επιβεβαίωση), optional room/responsible_doctor/hospital_mrn/status (εγκεκριμένο κείμενο ή null, όχι φανταστικές τιμές). Το inpatient bridge απαιτεί inpatient rows. Χωριστή εξωτερική πηγή πραγματικών PATIENTS και OUTPATIENT_SCHEDULE, με approved PatientType, πραγματική θεραπεία/ώρα/ημέρες/θεραπευτή, χωρίς άγνωστα populated metadata ή duplicate IDs. Πρώτα ολοκλήρωσε και επιβεβαίωσε τη δομή/πηγή και το καθαρό VBA shell, τελευταίο το provider schedule JSON `diomil-private-provider-schedule-v1` με hashes τελικών inputs, `availability`/ `shared_slots` όπως στο runbook. Ούτε clinical JSON ούτε provider JSON να χρησιμοποιούνται σαν auto-generated κλινική έγκριση. Current-program και approval sign-off εκτός repo.

### Γ. Σειρά ενεργειών όταν επιστρέψει ρητά στο PC

- Μία ασφαλής fast-forward ενημέρωση branch, μόνο με καθαρό tracked state, διατήρηση τοπικής εργασίας/αρχείων. Εάν έχει ήδη συγχρονιστεί, δεν χρειάζεται ξανά. Καταγεγραμμένη πλήρης green CI στο source commit 94db7c4· δεν απαιτείται επανάληψη όλων των historical tests για να αρχίσει η ιδιωτική προετοιμασία. Αν το τοπικό περιβάλλον/κώδικας έχει διαφορετική έκδοση ή σφάλμα, τότε μόνο στοχευμένος έλεγχος.
- Επιλογή εγκεκριμένων inputs Α, διαδικασία Β, ένα read-only checklist αν χρειάζεται. Για την ενημερωμένη εκτελέσιμη ακολουθία δες **την πρώτη ενότητα του υπάρχοντος `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`**, όχι τα archived commands. Standalone diagnostics μόνο όταν απαντούν σε συγκεκριμένο blocker, όχι υποχρεωτικά επανάληψη όλων.
- Μία πλήρης read-only εντολή `prepare_real_data_practice_demo.py` με **όλα τα έξι inputs**, χωρίς --build. Μόνο exit 0 ΚΑΙ explicit READ-ONLY PREFLIGHT GREEN ΚΑΙ ανεξάρτητη πραγματική έγκριση επιτρέπουν το επόμενο στάδιο. PARTIAL, missing files ή safety stop = όχι build. Δεν αποδέχεται unconfirmed προτάσεις με αλλαγή filename.
- Με τα ίδια αμετάβλητα inputs, ρητή έναρξη δεύτερου σταδίου, νέο μοναδικό private *_DEMO_*.xlsm, native COM save/close/reopen και independent readback, ποτέ overwrite. Freeze οποιοδήποτε πραγματικό failed/uncertain artifact και επόμενη διάγνωση στο συγκεκριμένο blocker. Το αρχικό manual χρησιμοποιείται παράλληλα και παραμένει authoritative.
- Τρίτο στάδιο: πλήρης ΔΕ–ΠΑ 1:1 private parity όλων των θεραπευτικών εγγραφών/κλινικών facts/εξωτερικών, γνωστή μορφοποίηση MASTER (ώρα πάνω, ημέρες κάτω, όπου προβλέπεται provider), μηδέν PM initial appointments και ξεχωριστή source namespace. Πραγματικές ενέργειες patient registration/edit, recurring changes, approved shared slots, daily absence/replacement, temporary exit/return, medical deferral, final discharge/transfer, staff/student lifecycle, daily close, restore, history/stats/management και protection/reopen. Να ελέγχονται ακριβείς αλλαγές και μη επηρεαζόμενα records· όχι επαναφορά αρχείου με overwrite. Τα ακόμη read-only ή pure PM edit/conflict/medical-scope εργαλεία δεν αποτελούν writable native workflow και χρειάζονται ολοκλήρωση/επαλήθευση όπου λείπουν πριν χαρακτηριστεί πλήρες. Αν κάτι δεν υποστηρίζεται ακόμη από την εφαρμογή, καταγράφεται ως μη υλοποιημένο, όχι operator pass. DEMO candidate verified ≠ πλήρης αποδοχή/production.

**STOP τώρα.** Ολοκληρώθηκε η συγκεντρωτική τεχνική προετοιμασία οδηγιών. Δεν εγκρίθηκαν πραγματικοί ασθενείς, δεν υπάρχει ακόμη νέο πραγματικό XLSM, δεν δωρίσαμε τεχνητά κλινικά δεδομένα. Ο επόμενος operator κύκλος είναι PC-only όταν ενημερώσει. Νέα cloud source εργασία μόνο για συγκεκριμένη ουσιαστική λειτουργικότητα, όχι εκ νέου αναζήτηση/επαναλαμβανόμενα tests χωρίς λόγο.

## 4.17 Επιστροφή χειριστή στο PC — 17/11/2026

**CURRENT RESUME:** Ο χρήστης δήλωσε ρητά «Επέστρεψα στο PC». Η προσωρινή παύση του §4.6 έληξε. Αυτό επιτρέπει τις τοπικές ενέργειες προετοιμασίας, ΟΧΙ αυτόματη κλινική έγκριση ή native build. Τελευταίο επαληθευμένο pre-return HEAD `449032a5aa6d73108598dfca8203402e2715d8cf`, Actions [44103624549](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/44103624549) 4/4 GREEN. Τελευταίο code checkpoint `94db7c4`, 2.113 pytest ανά OS (§4.15). Οι §§4.13–4.16 υπερισχύουν παλαιότερων οδηγιών (ιδίως blanket F/U/V stop και παλαιά πενταμερή inputs). Καμία σημερινή τοπική εκτέλεση δεν έχει ακόμη επιβεβαιωθεί.

**Επόμενη μία ενέργεια:** ασφαλής συγχρονισμός του local branch με clean tracked state, `git fetch origin` και `git merge --ff-only` μόνο, χωρίς force/reset/clean. Κατόπιν, επιλογή του εγκεκριμένου τρέχοντος programme XLSX και του αντίστοιχου proposal review (ή ιστορικού ζεύγους μόνο για προετοιμασία, όχι ως αυτόματα τρέχον), και μία εκτέλεση `prepare_trial_review_checklist.py --programme ... --review ... --output ...` σε νέο ιδιωτικό αρχείο εκτός Git, εφόσον δεν υπάρχει ήδη checklist για το ίδιο αμετάβλητο ζεύγος. Χρησιμοποίησε file picker, όχι αναζήτηση ακριβούς ονόματος. Το checklist περιέχει μόνο locators και generic reasons, καθώς και private input hashes· δεν είναι approved review. Στείλε μόνο aggregate CLI counts/exit code, ποτέ workbook/ονόματα/IDs/ιδιωτικές διαδρομές/hashes. Έπειτα ιδιωτική review/clinical/provider/source προετοιμασία κατά §4.16. Μην ξανατρέξεις τα 2.113 tests χωρίς συγκεκριμένο τοπικό πρόβλημα. Μην εκτελέσεις full preflight πριν υπάρξουν έξι εγκεκριμένα inputs· μη χρησιμοποιήσεις --build σε αυτή τη φάση. Το πλήρες native real DEMO και η λειτουργική αποδοχή εξακολουθούν να εκκρεμούν.

## 4.18 Τοπικός συγχρονισμός επιβεβαιώθηκε — 17/11/2026

Ο χειριστής επέστρεψε το πραγματικό αποτέλεσμα `COMMIT: 7d15049` και `CHECKLIST TOOL: True`. Ο ασφαλής συγχρονισμός και η παρουσία του `Python/tools/prepare_trial_review_checklist.py` έχουν αποδειχθεί στο PC. **Δεν επαναλαμβάνουμε Git pull ούτε τα 2.113 tests για να τρέξει το ήδη διαθέσιμο εργαλείο.** Επόμενο: μία ιδιωτική δημιουργία companion checklist με δύο Windows file pickers (εγκεκριμένο programme και αντίστοιχο review, ιστορικό ζεύγος μόνο ως προετοιμασία), νέο αποκλειστικό output στον εγκεκριμένο ιδιωτικό φάκελο του review με timestamp. Η ακύρωση σταματά ολόκληρο το block. Το εργαλείο ελέγχει distinct XLSX/εκτός Git/source hashes και independent readback. Στέλνονται μόνο κατηγορίες/aggregate counts/exit code. Δεν αποτελεί approved review, δεν αλλάζει πηγές, δεν έχει ακόμη εκτελεστεί σε αυτό το checkpoint. Εάν υπάρχει ήδη checklist για τα ίδια ακριβώς snapshots, χρησιμοποιείται εκείνο. Το τρέχον πρόγραμμα και οι έξι εγκεκριμένες πηγές εξακολουθούν να εκκρεμούν. Καμία native δημιουργία ακόμη.

## 4.19 Ιδιωτική λίστα εκκρεμοτήτων, ανεξάρτητη ανάκτηση — 17/11/2026

Ο τοπικός χειριστής ανέφερε `SAFETY STOP: checklist failed (TrialReviewChecklistError). EXIT_CODE=2`. Η CLI εμφανίζει μόνο την κλάση, άρα η ακριβής τοπική αιτία **δεν αποδείχθηκε** και η τοπική εκτέλεση δεν χαρακτηρίζεται GREEN. Τα δύο πραγματικά, ήδη ανεβασμένα XLSX βρέθηκαν στην ιδιωτική συνομιλία/runtime. Ελέγχθηκαν τα απαιτούμενα headers· συμφωνούν. Το review δεν έχει worksheet XML dimension metadata, ενώ περιέχει τα δεδομένα του· αυτό ΔΕΝ αποδεικνύεται αιτία του local stop. Ανεξάρτητη ιδιωτική ανάγνωση OOXML και εφαρμογή των ίδιων κανόνων inventory (όχι εκτέλεση του repository CLI) παρήγαγε νέο `DIOMIL_CHECKLIST_20261117.xlsx` με 271 task rows. Δημιουργία XLSX με artifact_tool, επαναφόρτωση και ακριβής επαλήθευση όλων των task rows/headers/μεταδεδομένων, total 271, hashes και των δύο αρχικών unchanged. Το αρχείο δίνεται ιδιωτικά στον χειριστή, **δεν** προστίθεται στο Git, περιέχει locators/γενικούς λόγους και private source digests, όχι clinical values. Counts: IDENTITY 106, HOLD 61, CONTEXT 30, F_RECLINING 5, U_AFTERNOON 19, REPORTED_OVERLAP 11, SAME_SPECIALTY 4, CROSS_SPECIALTY 11, WEEKLY 22 (NON_PROPOSED_STATE), SOURCE_LINK 2 (TREATMENT_ROW_NOT_IN_REVIEW), V_POOL 0. Οι κατηγορίες επικαλύπτονται. Ο αριθμός 271 δεν είναι μοναδικοί ασθενείς ή επιβεβαιωμένα λάθη. Καμία ταυτότητα/θεραπεία δεν εγκρίθηκε ούτε διορθώθηκε. Η πηγή 10/11 παραμένει ιστορική έως ρητή έγκριση επικαιρότητας.

**Επόμενο:** χρησιμοποιούμε τη δημιουργημένη ιδιωτική checklist για review της συγκεκριμένης πηγής, όχι άλλη ίδια επιλογή/αναδημιουργία από τον χειριστή. Να παραμένουν εκτός Git τα workbook/hash/ονόματα/IDs. Αν αλλάξει source ή απαιτείται εκ νέου το συγκεκριμένο CLI, τότε στοχευμένη διάγνωση με allowlisted μη ταυτοποιητικούς reason codes και privacy-preserving tests, όχι blanket bypass. Οι 6 εγκεκριμένες πηγές του §4.16, current-program sign-off, full native preflight, build και operator acceptance εξακολουθούν να εκκρεμούν.

## 4.20 Χρήση ιδιωτικής checklist — διευκρίνιση εκκρεμοτήτων

Rechecked the actual private checklist against both already uploaded inputs: its two stored SHA bindings match the input bytes. All 22 WEEKLY rows are `HOLD_CONFLICT` and are already members of the independent collision groups; do not ask the operator to resolve those a second time. The 2 SOURCE_LINK tasks refer to source rows with populated C:V fields but no explicit B value; do not infer a patient identity or classify them as actual patients without checking the source layout/context. This is read-only evidence, not clinical resolution. Operator keeps the checklist privately as a reference; no need to manually complete 271 rows. First establish whether the historical 10/11 snapshot is the intended isolated trial basis or whether a newer matched pair is required. Reuse actual evidence and group overlapping tasks; ask only the essential clinical decisions. No input modifications, approvals or native build.

## 4.21 Επιλογή πραγματικού αρχικού snapshot και σταδιακή ενημέρωση — 17/11/2026

**ΑΠΟΦΑΣΗ ΧΡΗΣΤΗ:** Το ήδη εξετασμένο A:V πρόγραμμα της 10/11/2026 είναι η επιλεγμένη αρχική βάση του απομονωμένου πραγματικού DEMO. Δεν ζητάμε νεότερα προγράμματα ή νέο χωρισμό τους σε στήλες. Ο χρήστης θέλει να ενημερώνει και να διορθώνει σταδιακά το πρόγραμμα στη δουλειά, μέσα στο λειτουργικό DEMO, χωρίς να χρειάζεται πλήρης επικαιροποίηση του ιστορικού snapshot εκ των προτέρων. Αυτή η απόφαση επιλύει την εκκρεμότητα **επιλογής snapshot για το trial**, όχι την τρέχουσα κλινική εγκυρότητα, ατομικές ταυτοποιήσεις, F/V/U, διαθέσιμες ώρες, shared assignments ή την αποδοχή λειτουργικής εισαγωγής. Διατηρείται το πρωτότυπο. Δεν χρησιμοποιείται το ιστορικό DEMO σαν τρέχον πρόγραμμα περίθαλψης.

Νέος ανεξάρτητος read-only έλεγχος των ίδιων ήδη διαθέσιμων XLSX: οι 61 hold rows ταξινομούνται σε 29 NO_TIME, 11 NO_EXPLICIT_DAY, 16 UNRECOGNIZED_TIME, 4 UNRECOGNIZED_DAY, 1 DUPLICATE_DAY. Οι 764 weekly rows έχουν 742 PROPOSED_EXACT_SOURCE και 22 HOLD_CONFLICT. Τα 2 SOURCE_LINK αφορούν πραγματικά populated C:V χωρίς B: μία σειρά με H:I/J:K:L και μία με G, όχι αυτομάτως νέους ασθενείς. Δεν αλλάχτηκε κανένα αρχείο. Επόμενη τεχνική εργασία: εφάρμοσε/έλεγξε τον ήδη συμφωνημένο κανόνα populated clock + blank weekday = ΔΕ–ΠΑ στα 11 σχετικά holds εφόσον είναι κανονική θεραπεία και όχι ειδικό πεδίο, με ακριβή πηγή και χωρίς αυτόματη έγκριση. Ξεχώρισε τα υπόλοιπα τεχνικά parse προβλήματα από πραγματικές αμφισημίες. Μην αναθέσεις τον χειριστή να συμπληρώσει 271 tasks ούτε ζήτησε επαναληπτικό audit. Τα 6 inputs και οι fail-closed rules της πλήρους εισαγωγής εξακολουθούν να ισχύουν. Αν χρειάζεται μερική, διορθώσιμη πρώτη δοκιμή, σχεδίασε ρητά χωριστό scope, με πλήρες ledger των μη εισαγόμενων κρατήσεων και πραγματικά επιβεβαιωμένη επιλεξιμότητα αυτών που εισάγονται· δεν υποστηρίζεται ακόμη από τον full builder. Μην χαλαρώσεις full gate ή παρουσιάσεις pending entries ως ενεργές θεραπείες. Καμία native εισαγωγή/κλινική έγκριση σε αυτό το checkpoint.

## 4.22 Κριτήριο έναρξης πραγματικής παράλληλης δοκιμής — 17/11/2026

Ο χρήστης διευκρίνισε: «όταν συγχρονιστεί με το τρέχον της εργασίας να μπορώ να το δοκιμάζω». Αρχική βάση του απομονωμένου DEMO παραμένει το 10/11 (§4.21). Μετά την ασφαλή εισαγωγή/δημιουργία του ιδιωτικού XLSM, ο χειριστής θα το επικαιροποιεί σταδιακά με τις ήδη υποστηριζόμενες λειτουργίες (εισαγωγή/έξοδος, αλλαγές recurring προγράμματος, θεραπευτές και ώρες), όχι με αυθαίρετες αλλαγές στα κρυφά source sheets. Η κατάσταση «έτοιμο για πραγματική παράλληλη δοκιμή» απαιτεί συμφωνία των ενεργών ασθενών, εβδομαδιαίων αναθέσεων ΔΕ–ΠΑ, σχετικών πραγματικών capacity/availability και ημερήσιων εξαιρέσεων με το τρέχον πρόγραμμα εργασίας για τη συγκρινόμενη ημέρα. Δεν θεωρούμε το ιστορικό snapshot τρέχουσα πραγματικότητα. Τυχόν μη υλοποιημένες διαδρομές ή αποκλίσεις μένουν ρητά εκκρεμείς, δεν αποκρύπτονται. Η παραγωγική τρέχουσα διαδικασία παραμένει authoritative· το DEMO είναι ανεξάρτητο, για παράλληλη σύγκριση συμπεριφοράς, όχι ανάληψη πραγματικών κλινικών αποφάσεων. Δεν έχει ακόμη δημιουργηθεί πραγματικό XLSM ούτε αποδειχθεί ο συγχρονισμός με το πρόγραμμα της εργασίας. Επόμενο engineering: §4.21, τεχνικός διαχωρισμός επιλύσιμων source holds από αμφίσημα/κλινικά, χωρίς νέο κύκλο γενικών audits.

## 4.23 Έναρξη πραγματικού UI review και τεχνική αναθεώρηση επιλεγμένου snapshot — 17/11/2026

Ο χρήστης ζήτησε να ξεκινήσει η δοκιμή πραγματικής λειτουργίας, σχεδιασμού, θέσης κουμπιών κλπ. **Μην αναβάλλεις το UI feedback μέχρι την πλήρη clinical migration.** Προαιρετική συγκεκριμένη πρώτη οθόνη: αποκλειστικό byte-exact ιδιωτικό αντίγραφο του ήδη γνωστού native `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` με `Python/tools/prepare_native_probe_copy.py --source ... --output ...`, σε νέο τοπικό φάκελο εκτός Git. Μόνο οπτικός/πλοηγικός έλεγχος MASTER και υπάρχουσας σχεδίασης, όχι real-data DEMO, όχι καθαρό shell, όχι απόδειξη ότι όλα τα νεότερα modules έχουν ενσωματωθεί. Frozen CC αμετάβλητο. Μην ζητάς screenshots με patient-identifying στοιχεία ή επανάληψη όλων των historical tests· ζητάμε συγκεκριμένα UI observations, ιδίως θέση/ορατότητα κουμπιών, παλαιά μορφοποίηση MASTER και πλοήγηση. Native copy/open δεν έχει ακόμη πραγματοποιηθεί στο PC και καμία κλινική ενέργεια δεν εξουσιοδοτείται από αυτό.

**Πραγματική τεχνική πρόοδος ιδιωτικής προετοιμασίας:** Ανεξάρτητη OOXML ανάγνωση των ήδη ανεβασμένων 10/11 programme + proposal review: όλα τα 11 `NO_EXPLICIT_DAY` (10 Εργοθεραπεία, 1 Ψυχολογία) έχουν όντως έγκυρη αρχική ώρα, κενή ημέρα, συνδεδεμένη ταυτότητα source row και υπάρχον optional provider. Με τον ήδη εγκεκριμένο κανόνα ώρα+κενή ημέρα=ΔΕ–ΠΑ δημιουργήθηκαν 55 τεχνικές προτάσεις, 0 νέες same-patient exact-start συγκρούσεις είτε μεταξύ τους είτε με τις 764 υπάρχουσες. Έξι από αυτές συμμετέχουν σε κοινό provider/day/time group βάσει canonical provider normalization, και θέλουν exact membership approval. Οι 29 NO_TIME (Ψυχολογία) είναι όντως blank clocks με day/provider metadata· δεν δημιουργήθηκε καμία συνεδρία, δεν διαγράφηκαν, χρειάζονται ρητή disposition. Οι 16 UNRECOGNIZED_TIME παραμένουν μη αναγνώσιμες με τον αυστηρό clock parser, 4 UNRECOGNIZED_DAY και 1 DUPLICATE_DAY ανεπίλυτα. Καμία επέμβαση στα 15 παλαιά conflict groups.

Δημιουργήθηκε ιδιωτικό `DIOMIL_REVIEW_TECHNICAL_DRAFT_20261117.xlsx` (artifact_tool, νέο αρχείο εκτός Git), με 7 φύλλα: αρχική δομή review, 819 weekly proposal rows (764 original exact preserved +55), 50 active hold rows (οι 11 day proposals μεταφέρθηκαν στο ξεχωριστό audit `ΤΕΧΝΙΚΕΣ_ΠΡΟΤΑΣΕΙΣ` με original rows και αιτιολογία, όχι σιωπηλή διαγραφή), 11 παλιές overlap rows, νέο `ΚΟΙΝΑ_SLOTS_ΠΡΟΣ_ΕΛΕΓΧΟ` με 6 προτάσεις. Οι 106 patient review rows και E/F/G/H values είναι ακριβώς αμετάβλητα, 0 confirmed IDs. Independent exact value readback και για τα 7 φύλλα, 819/50/6, πρωτότυπα SHA unchanged. Το draft ΔΕΝ είναι εγκεκριμένο review, δεν έχει περάσει full CLI και δεν είναι native demo. Μη δημοσιεύσεις κανένα ιδιωτικό αρχείο/τιμές/hashes στο GitHub. Προχωράμε τεχνικά από αυτό το draft, χωρίς να ζητάμε νέο source inventory. Για πραγματική εισαγωγή εξακολουθούν να απαιτούνται εγκεκριμένες ταυτότητες/clinical fields, αποφάσεις των άλλων holds και conflicts, F/V/U, outpatient/provider source, καθαρό shell, full preflight, native COM και πραγματική αποδοχή. Οι ενέργειες UI copy είναι ανεξάρτητος πρόωρος οπτικός έλεγχος, όχι υποκατάστατο του ζητούμενου πραγματικού demo.

## 4.24 Πρώτη παρατήρηση χειριστή στο πραγματικό Excel UI — 17/11/2026

Ο χρήστης άνοιξε το απομονωμένο UI αντίγραφο της προγενέστερης CC έκδοσης και αναφέρει ότι όλα μοιάζουν στη θέση τους. Αυτό είναι θετικό **οπτικό feedback**, όχι ολοκληρωμένο acceptance κουμπιών/λειτουργιών. Δύο συγκεκριμένα ευρήματα MASTER: τα κελιά «Θάλαμος» και «Ασθενής» έγιναν κίτρινα όπως η επισήμανση λοιμωδών (δεν έχει ακόμη αποσαφηνιστεί αν αφορά headers, όλη τη στήλη ή συγκεκριμένες γραμμές)· δεν εμφανίζονται οι ιατροί, πιθανώς λόγω παλαιότερης έκδοσης ή απουσίας δεδομένων. Source inspection: `master_visual_style.py` έχει χωριστά identity header/body χρώματα και εφαρμόζει κίτρινο B:C μόνο βάσει infectious flag· `master_projection_writer.py` αντιγράφει μορφοποίηση B:C από την πρώτη γραμμή template και μετά εφαρμόζει κίτρινο στους λοιμώδεις, πιθανή κληρονόμηση χρωματισμού εάν το template είναι ήδη κίτρινο. Αυτό είναι συγκεκριμένος υποψήφιος μηχανισμός, **όχι αποδεδειγμένη αιτία** της παρατήρησης. Επίσης `_render_patient_identity` εμφανίζει `ResponsibleDoctor` όταν υπάρχει, δεύτερη γραμμή με κόκκινο σταυρό, και το schema επιτρέπει legacy χωρίς doctor. Να εξακριβωθεί η πραγματική native συμπεριφορά και στα δύο κριτήρια στο νέο candidate: ουδέτερο identity header/noninfectious σώμα, κίτρινη επισήμανση μόνο κατά τον κανόνα λοιμωδών, σωστός εγκεκριμένος ιατρός στην καθιερωμένη δεύτερη γραμμή, χωρίς αυθαίρετη κλινική συμπλήρωση ή νέα διάταξη. Καμία διόρθωση VBA/Excel δεν εκτελέστηκε τώρα. Ο χειριστής κρατά το `DIOMIL_REVIEW_TECHNICAL_DRAFT_20261117.xlsx` ιδιωτικά μαζί με τα δύο αρχικά αρχεία, δεν είναι το λειτουργικό DEMO, δεν απαιτεί χειροκίνητη συμπλήρωση τώρα. Επόμενο engineering: προετοιμασία από το επιλεγμένο snapshot και scope πραγματικών κλινικών εγκρίσεων, όχι νέα γενική UI ανακατασκευή ή επαναληπτικά diagnostics. Το πλήρες native πραγματικό DEMO παραμένει PENDING.

## 4.25 Επιβεβαίωση μοναδικού οδηγού και συνέχεια — 17/11/2026

Ρητή επαναβεβαίωση χρήστη: ο παρών οδηγός διαβάζεται στην έναρξη κάθε συνομιλίας και συμπληρώνεται συνεχώς και ανελλιπώς, στην ίδια εργασία με κάθε ουσιαστική αλλαγή, απόφαση, απόδειξη, αποτυχία και ακριβές επόμενο βήμα. Ενεργός κλάδος `work/afternoon-medical-scope-20261115`, όχι ο ιστορικός feature branch. Οι νεότερες §§4.21–4.24 υπερισχύουν των παλαιών resume σημείων. Στη νέα συνομιλία δεν ανακτήθηκε το ιδιωτικό technical draft μέσω αναζήτησης αρχείων· αυτό δεν αναιρεί την προηγούμενη δημιουργία/επαλήθευσή του. Επόμενο: συνέχιση από το ίδιο ιδιωτικό draft όταν καταστεί προσβάσιμο, παράλληλα με ασφαλή στοχευμένη source εργασία όπου είναι εφικτή. Δεν έγινε τώρα νέα επεξεργασία ιδιωτικών δεδομένων, αλλαγή κώδικα ή native δοκιμή. Καμία επανάληψη source inventory, κανένα νέο synthetic DEMO, καμία δημοσίευση ιδιωτικών δεδομένων. Πραγματικό native DEMO PENDING.

## 4.26 Επαναφορά πρόσβασης στα τέσσερα ιδιωτικά αρχεία — 17/11/2026

Ο χειριστής ανέβασε ξανά το επιλεγμένο αρχικό programme, το αρχικό identity proposal, το companion checklist και το technical draft. Και τα τέσσερα είναι προσβάσιμα στο παρόν runtime: επιτυχής ανάγνωση OOXML workbook metadata και ZIP integrity, αντίστοιχα 74, 5, 2 και 7 φύλλα. Αυτό επιλύει το εμπόδιο πρόσβασης της §4.25· δεν είναι νέος έλεγχος κλινικού περιεχομένου ή έγκριση. Δεν άλλαξε κανένα workbook. Επόμενο: στοχευμένη τεχνική επεξεργασία των υπολειπόμενων holds από το υπάρχον draft με αναφορά στο αρχικό programme, διαχωρισμός ασφαλώς αναγνώσιμων τιμών από πραγματικές αμφισημίες· διατήρηση όλων των αρχικών και μη εγκεκριμένων εγγραφών. Όχι νέο γενικό audit και όχι νέα αίτηση αποστολής όσο τα αρχεία είναι προσβάσιμα. Πραγματικό native DEMO παραμένει PENDING.

## 4.27 Στοχευμένη ταξινόμηση 50 ενεργών holds — 17/11/2026

Στο παρόν runtime επαληθεύτηκε πράγματι ZIP integrity και workbook metadata των τεσσάρων ιδιωτικών XLSX (74/5/2/7 φύλλα). Στο υπάρχον technical draft διαβάστηκε το φύλλο ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ: 50 data rows, 29 NO_TIME, 16 UNRECOGNIZED_TIME, 4 UNRECOGNIZED_DAY, 1 DUPLICATE_DAY. Από τις 16 ώρες, 7 έχουν σαφή αριθμητική μορφή ώρας με ελληνικό τόνο ως διαχωριστικό, 8 είναι μόνο αστερίσκος (χωρίς ώρα), 1 περιέχει μη αριθμητικούς χαρακτήρες και παραμένει αμφίσημη. Σε 4 day holds υπάρχουν μη canonical συντομογραφίες ή πιθανή μετατόπιση πεδίων· το duplicate-day hold έχει δύο ώρες και επαναλαμβανόμενη ημέρα. Καμία αυτόματη κλινική ερμηνεία. Επόμενο ασφαλές engineering: εξέταση του υπάρχοντος source clock parser για στενή, δοκιμασμένη αναγνώριση του συγκεκριμένου διαχωριστικού, χωρίς γενική χαλάρωση, και έλεγχος των 7 έναντι original cell/provenance και συγκρούσεων πριν προστεθούν ως τεχνικές (όχι εγκεκριμένες) προτάσεις. Οι 8 αστερίσκοι, 29 blank clocks και λοιπές αμφισημίες παραμένουν στο ledger. Δεν δημιουργήθηκαν νέες κρατήσεις, δεν τροποποιήθηκαν αρχεία ούτε εκτελέστηκαν tests σε αυτό το στάδιο. Καμία ιδιωτική τιμή/ταυτότητα στο GitHub. Native DEMO PENDING.

## 4.28 Επαναχρησιμοποίηση υπάρχοντος μηχανισμού ώρας — 17/11/2026

Ο χρήστης επισήμανε ότι υπάρχει ήδη εργαλείο για τις διαφορετικές γραφές ώρας. Επιβεβαιώθηκε με ανάγνωση του ενεργού branch ότι `Python/rehab_core/programme_mapping.py::parse_clock` είναι ο κοινός parser που καλούν `split_weekly_cell`, `trial_source_weekday_projection.py` (μέσω `make_weekly_entries` και απευθείας F/V) και `weekly_slots_source.py`. Υποστηρίζει native Excel time/fractions, τελεία, άνω-κάτω τελεία, κόμμα, όχι ελληνικό τόνο ως separator. Δεν δημιουργούμε δεύτερο εργαλείο. Πριν από scoped διόρθωση ελέγχουμε τυχόν άλλους υφιστάμενους normalizers και tests ώστε να μην διπλασιαστεί λογική ή αλλάξει αποδεκτή συμπεριφορά. Η αναζήτηση default branch δεν επέστρεψε αποτελέσματα, άρα δεν έχει αποκλειστεί άλλος normalizer. Καμία αλλαγή κώδικα, ιδιωτικού draft ή νέα εκτέλεση tests σε αυτό το βήμα. Επόμενο: επαναχρησιμοποίηση/στενή επέκταση του υπάρχοντος μηχανισμού με regression tests και έλεγχο πηγαίας αντιστοίχισης. Native DEMO PENDING.

## 4.29 Ρητός κανόνας αυτόνομης συνέχειας — 17/11/2026

Ο χρήστης επαναβεβαίωσε: προχωράμε χωρίς να περιμένουμε «ΟΚ» ή νέα εντολή ανά ενδιάμεσο στάδιο. Ζητάμε παρέμβαση μόνο για αναγκαία ενέργεια στο τοπικό PC, πραγματική απόφαση του χειριστή ή αντικειμενικό blocker. Ενημερώνουμε τον μοναδικό οδηγό σε κάθε ουσιαστικό στάδιο, επαναχρησιμοποιούμε υπάρχοντα εργαλεία και αποφεύγουμε επαναληπτικές εργασίες. Αμέσως επόμενο τεχνικό βήμα §4.28: υφιστάμενος parser και συναφή tests, στενή διόρθωση μόνο εφόσον χρειάζεται. Η οδηγία αφορά την αυτόνομη εργασία κατά την ενεργή συνομιλία, όχι υπόσχεση εργασίας στο παρασκήνιο.

## 4.30 Στενή επέκταση υφιστάμενου parser ώρας — 17/11/2026

Επιβεβαιώθηκε με πραγματική OOXML read-only ανάγνωση του technical draft ότι και οι 7 υπό εξέταση εγγραφές έχουν διαχωριστικό U+0384, ώρα 1 ή 2 ψηφίων και λεπτά 2 ψηφίων. Και τα 4 ιδιωτικά XLSX πέρασαν ZIP integrity. Commit `c1a2fb0`: ο υπάρχων `rehab_core/programme_mapping.py::parse_clock` αναγνωρίζει μόνο την ακριβή μορφή 1–2 ASCII ψηφία + U+0384 + 2 ASCII ψηφία, μετατρέποντάς την σε colon format και επαναχρησιμοποιώντας τον έλεγχο ορίων. Commit `44c6fd6`: 11 parametrized synthetic cases στο υπάρχον `Tests/test_trial_source_weekday_projection.py`, 5 θετικές/6 αρνητικές, με διατήρηση ημερών/provider/source cell και fail-closed αμφισημίες. Δεν δημιουργήθηκε δεύτερος parser. Δεν έχει επιβεβαιωθεί εκτέλεση CI ή pytest για τα νέα commits, επομένως όχι GREEN. Δεν έχει ακόμη τροποποιηθεί το ιδιωτικό draft, ούτε εγκρίθηκε ή εισήχθη θεραπεία. Επόμενο: επιβεβαίωση test αποτελεσμάτων και έλεγχος των 7 έναντι ακριβούς original source/provenance/overlap και shared-slot membership πριν από νέα ιδιωτική αναθεώρηση. Καμία ιδιωτική τιμή στο δημόσιο repo. Native DEMO PENDING.

## 4.31 Πραγματική επαλήθευση επτά clock holds — 17/11/2026

Στο παρόν runtime έγινε στοχευμένη ανεξάρτητη OOXML read-only ανάγνωση του αρχικού programme και του technical draft. Και τα 7 U+0384 holds είναι Λογοθεραπεία, J/K/L, συμφωνούν στο source row/name και αρχική ώρα, έχουν canonical ημέρες και provider. Η ακριβής αριθμητική ερμηνεία δίνει 8 εβδομαδιαίες υποψήφιες συνεδρίες (μία εγγραφή δύο ημερών). Από αυτές 6 δεν έχουν νέα same-patient exact-start σύγκρουση, 2 συγκρούονται με υπάρχουσες προτάσεις. Δεν βρέθηκε νέα provider/day/exact-start σύγκρουση με ανεξάρτητη accent/case/punctuation normalization· όχι απόδειξη duration/capacity ή ταυτότητας provider. Δεν τροποποιήθηκε workbook και δεν εκτελέστηκε ακόμη πλήρης parser regression/CI. Επόμενο: ιδιωτική αναθεώρηση με διατήρηση source provenance, καταγραφή και των 8 ως τεχνικών υποψηφίων, 2 συγκρούσεις ρητά held χωρίς αυθαίρετη μετακίνηση, καμία αυτόματη clinical approval. Native DEMO PENDING.

## 4.32 Πραγματικό ιδιωτικό technical draft R2 — 17/11/2026

Δημιουργήθηκε και επαληθεύτηκε με artifact_tool και ανεξάρτητη OOXML readback το ιδιωτικό `DIOMIL_REVIEW_TECHNICAL_DRAFT_20261117_R2.xlsx`. Οι επτά αρχικές clock holds συνδέθηκαν με την ακριβή J/K/L πηγή και μεταφέρθηκαν στο audit. Προστέθηκαν 8 τεχνικές weekly προτάσεις: 6 χωρίς νέα same-patient exact-start σύγκρουση, 2 με σύγκρουση· σημάνθηκαν HOLD_CONFLICT και οι 2 αντίστοιχες προϋπάρχουσες εγγραφές. Σύνολο 827 weekly, 801 PROPOSED_EXACT_SOURCE, 26 HOLD_CONFLICT, 43 active source holds, 13 καταγεγραμμένες overlap rows, 18 audit rows, 6 προηγούμενα shared-slot review rows αμετάβλητα. Όλες οι 106 identity rows ακριβώς αμετάβλητες, καμία κλινική έγκριση. Και τα 4 inputs με unchanged SHA, ZIP integrity του R2 green. Full CLI/CI/native acceptance δεν εκτελέστηκε σε αυτό το στάδιο. Το R2 παραδόθηκε ιδιωτικά, όχι στο GitHub. Επόμενο: εξέταση 43 υπολειπόμενων holds και ειδικά των 4 day ambiguities/1 duplicate day, διαχωρισμός ασφαλών τεχνικών διορθώσεων από αναγκαίες αποφάσεις. Μην σταματάς για ΟΚ μετά από κάθε υποβήμα. Native DEMO PENDING.

## 4.33 Σημείο αναγκαίας απόφασης χειριστή — 17/11/2026

Ο χειριστής επιβεβαίωσε ότι κατέβασε το R2. Στο παρόν runtime διαβάστηκαν οι 43 ενεργές εγγραφές του R2, χωρίς αλλαγή αρχείου: 29 blank clocks, 8 αστερίσκοι Εργοθεραπείας, 1 αμφίσημη ώρα Λογοθεραπείας, 4 αμφίσημες ημέρες (δύο μη canonical γραφές και δύο πιθανές μετατοπίσεις provider/day), 1 επαναλαμβανόμενη ημέρα μεταξύ δύο ωρών. Δεν εικάζουμε κλινικές τιμές. Πρώτη αναγκαία ερώτηση προς χειριστή: πραγματική σημασία αστερίσκου στην Εργοθεραπεία και εάν η Ψυχολογία χωρίς ώρα δηλώνει θεραπεία χωρίς σταθερό ραντεβού ή μη προγραμματισμένη θεραπεία. Καμία αλλαγή workbook/parser, approval ή native build σε αυτό το στάδιο. Τα υπόλοιπα αμφίσημα rows παραμένουν προς επίλυση, όχι σιωπηλή παράλειψη.

## 4.34 Αποφάσεις χειριστή για αστερίσκο και Ψυχολογία — 17/11/2026

Ρητή απόφαση: ο αστερίσκος στην Εργοθεραπεία δεν αφορά το DIOMIL, αντιμετωπίζεται σαν να μην υπάρχει. Ψυχολογία χωρίς ώρα δηλώνει ουσιαστικά θεραπεία χωρίς σταθερό ωράριο στο snapshot· πιθανό νεότερο αρχείο ωραρίου θα μπορεί να προστεθεί αργότερα. Καμία δημιουργία timed slot χωρίς πραγματική ώρα. Μελλοντικά ώρες και ημέρες εισάγονται μόνο μέσω drop-down lists, όνομα προς το παρόν ελεύθερο κείμενο· ο χειριστής αναφέρει ότι στη ΦΘ και τις συναφείς θεραπείες δεν υπάρχουν τέτοια σημάδια. Η αγνόηση αστερίσκου είναι ειδική εγκεκριμένη Εργοθεραπεία legacy-source disposition, όχι γενική χαλάρωση parser. Να επαληθευτούν υπάρχοντα UI validation και import contracts πριν από αλλαγές.

Δημιουργήθηκε ιδιωτικό `DIOMIL_REVIEW_TECHNICAL_DRAFT_20261117_R3.xlsx` από R2 με artifact_tool. Οκτώ αστερίσκοι μεταφέρθηκαν από ενεργά holds στη νέα καρτέλα ΑΠΟΦΑΣΕΙΣ_ΧΕΙΡΙΣΤΗ, μαζί με 29 εγγραφές Ψυχολογίας που καταγράφονται ως εκκρεμές ωράριο χωρίς κράτηση και παραμένουν επίσης στα active holds μέχρι υποστήριξη της εισαγωγής/ωραρίου. Σύνολα 35 active holds (29 + 6), 37 ledger rows, 827 weekly proposals, 106 identities. Ανεξάρτητος OOXML έλεγχος ZIP integrity, exact nonempty cell values και των 5 ανέπαφων φύλλων (ΑΣΘΕΝΕΙΣ, ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ, ΕΠΙΚΑΛΥΨΕΙΣ, ΤΕΧΝΙΚΕΣ_ΠΡΟΤΑΣΕΙΣ, ΚΟΙΝΑ_SLOTS_ΠΡΟΣ_ΕΛΕΓΧΟ), hashes όλων των εισόδων unchanged. Καμία νέα κλινική έγκριση, καμία native/CI ή αλλαγή parser/UI σε αυτό το στάδιο. Το R3 δεν είναι full preflight approved: ο αυστηρός source gate χρειάζεται ρητή source-bound υποστήριξη των dispositions. Επόμενο: στοχευμένη εξέταση των 6 λοιπών αμφισημιών και επαναχρησιμοποίηση υπάρχοντος validation/untimed treatment contract, όχι επανάληψη γενικού audit. Το R3 παραδίδεται ιδιωτικά, όχι στο δημόσιο GitHub.

## 4.35 Στοχευμένη επαλήθευση έξι υπολειπόμενων αμφισημιών — 17/11/2026

Ο χειριστής κατέβασε το R3. Αναγνώστηκαν ιδιωτικά τα έξι non-NO_TIME holds του R3 και τα αντίστοιχα αρχικά A:V κελιά με OOXML read-only. Δύο ασαφείς ομαδοποιήσεις ημερών, δύο πεδία ημερών με κείμενο που μοιάζει με όνομα (το ένα με κενό provider, το άλλο με διαφορετικό populated provider), μία επαναλαμβανόμενη ημέρα σε δύο διαφορετικές ώρες και μία ώρα με μη αριθμητικούς χαρακτήρες. Ζητήθηκε συγκεντρωτική επιβεβαίωση των έξι πραγματικών αμφισημιών από τον χειριστή· όχι αυτόματη ερμηνεία. Δεν δημοσιεύονται πραγματικές τιμές ή source-row locators στο public repo. Καμία αλλαγή workbook, έγκριση ή εκτέλεση CI. Επόμενο: εφαρμογή μόνο των ρητών αποφάσεων σε νέα ιδιωτική αναθεώρηση, exact source/overlap/provider checks, και ασφαλής source-bound υποστήριξη legacy dispositions της §4.34. Το ανεξάρτητο ωράριο Ψυχολογίας παραμένει προσθέσιμο αργότερα. Native DEMO PENDING.

## 4.36 Επίλυση έξι αμφισημιών και ιδιωτικό R4 — 17/11/2026

Ρητές αποφάσεις χειριστή: αφαίρεση περιττού παρενθετικού χαρακτήρα από ημέρες, μεταφορά ονόματος από πεδίο ημερών στο πεδίο ονόματος με κενές ημέρες=ΔΕ–ΠΑ, αντιμετάθεση πεδίων ονόματος/ημερών, διόρθωση μη αριθμητικής ώρας και προσθήκη ελλείποντος διαχωριστικού ημερών. Στην ΕΦΑ επιτρέπονται διαφορετικές ώρες και δύο διακριτές συνεδρίες την ίδια ημέρα: δεν ενδιαφέρει το DIOMIL ο χαρακτηρισμός ατομικής/ομαδικής. Δεν επιτρέπεται αυθαίρετη γενίκευση σε ταυτόχρονα slots. Διατηρείται η αρχική πηγή, οι διορθώσεις είναι στο ιδιωτικό ledger.

Δημιουργήθηκε `DIOMIL_REVIEW_TECHNICAL_DRAFT_20261117_R4.xlsx` με artifact_tool από R3. Προστέθηκαν 19 εβδομαδιαίες προτάσεις, σύνολο 846 (818 PROPOSED_EXACT_SOURCE, 28 HOLD_CONFLICT). Μία νέα ομάδα ίδιας ώρας ασθενούς, τέσσερις πρόσθετες γραμμές για έλεγχο κοινών slots θεραπευτών με ανεξάρτητη κανονικοποίηση ονόματος. Οι έξι source holds μεταφέρθηκαν στο ledger αποφάσεων (43 συνολικά entries), μόνο 29 ενεργές εγγραφές χωρίς ωράριο Ψυχολογίας. Επαληθεύτηκαν μέσω independent OOXML readback οι 846 weekly rows, οι 106 identity rows ακριβώς αμετάβλητες, ledger counts, ZIP integrity και SHA unchanged όλων των input XLSX. Δεν έγινε πλήρες CLI/CI, καμία native εισαγωγή ή έγκριση ταυτοτήτων. Το R4 παραδόθηκε ιδιωτικά. Επόμενο: έλεγχος existing gate για source-bound corrections, legacy dispositions και διαφορετικά start times της ίδιας ημέρας (η τρέχουσα parse_weekday_groups απορρίπτει επαναληπτική ημέρα), χωρίς γενική χαλάρωση ή αυθαίρετη αλλαγή αρχικού programme. Ομαδοποίηση εκκρεμών patient/provider conflicts για χειριστή. Δεν απαιτείται να συμπληρώσει τώρα ωράρια Ψυχολογίας.

## 4.37 Υποστήριξη διακριτών ωρών ΕΦΑ την ίδια ημέρα — 17/11/2026

Ο χειριστής επιβεβαίωσε ότι κατέβασε και είδε το R4. Source commits `e48b317`, `9a15393`: ο υπάρχων programme_mapping parser επιτρέπει επανάληψη weekday μεταξύ ρητά χωρισμένων ομάδων μόνο για ΕΦΑ και μόνο εφόσον το weekday/start_time είναι μοναδικό. Ίδια ημέρα μέσα στην ίδια ομάδα ή ίδια ώρα σε επαναλαμβανόμενη ημέρα εξακολουθούν να απορρίπτονται. Ο weekly_slots_source reader χρησιμοποιεί επιπλέον την ώρα ως μέρος του source duplicate key μόνο στην ΕΦΑ. Άλλες ειδικότητες διατηρούν προηγούμενη συμπεριφορά. Commits `b76ce53`, `92c814c`: 6 νέες fictional regression cases, συμπεριλαμβανομένου weekly reader roundtrip. Δεν επαληθεύτηκε ακόμη η εκτέλεση pytest/CI, όχι GREEN. Επιθεωρήθηκε επίσης το `trial_demo_planner_mirror.py`: για διαφορετικές ώρες/παρόχους ή πολλαπλές συνεδρίες ίδιας ημέρας δεν παράγει legacy mirror, το granular view παραμένει διαθέσιμο. Αυτό χρειάζεται στοχευμένη native MASTER bulk-seed parity εξέταση χωρίς αλλαγή του ήδη αποδεκτού operational renderer. Η source reconciliation ακόμη διαβάζει raw A:V χωρίς υποστήριξη των εγκεκριμένων ιδιωτικών corrections/dispositions του R4. Επόμενο: source-bound ιδιωτική correction/disposition διαδρομή και verification των νέων tests, ακολούθως MASTER parity και ομαδοποιημένες πραγματικές συγκρούσεις. Καμία αλλαγή ιδιωτικού workbook ή native δοκιμή. Native DEMO PENDING.

## 4.38 Ρητές legacy dispositions χωρίς τεχνητά slots — 17/11/2026

Commits `42245da`, `429139c`, `fa214f1`: επέκταση υφιστάμενων preview_source_weekdays, trial_special_fields_audit και independent reconciliation. Ιδιωτική καρτέλα ΕΓΚΡΙΣΕΙΣ_ΧΩΡΙΣ_ΩΡΑ με SourceCell / ProgrammeSHA256 / Decision, μόνο ακριβής παλιός αστερίσκος Εργοθεραπείας (IGNORE_OCCUPATIONAL_MARK) ή κενή ώρα Ψυχολογίας με ημέρα/πάροχο (PSYCHOLOGY_TIME_PENDING). Έλεγχος ακριβούς αρχικού hash, literal τιμών, scope, duplicate/unsafe/stale entries· χωρίς δημιουργία slot ή γενική ανοχή. Οι άλλες θεραπείες του ασθενούς παραμένουν. Commits `18c1b57`, `e13236a`: 13 πρόσθετες synthetic περιπτώσεις tests. Η ανάκτηση workflow runs και combined status για e13236a δεν επέστρεψε αποτελέσματα, άρα εκτέλεση CI/pytest δεν έχει επιβεβαιωθεί. R4 αμετάβλητο, η νέα καρτέλα δεν έχει ακόμη συμπληρωθεί, 29 ενεργά holds παραμένουν. Δεν έχει υλοποιηθεί native αποθήκευση/προβολή εκκρεμούς Ψυχολογίας στο MASTER ούτε η διαδρομή των 5 διορθωμένων source πεδίων. Επόμενο: ολοκλήρωση/επαλήθευση διορθωμένης πηγής και legacy decision ledger με διατήρηση πρωτοτύπου, ομαδοποίηση 28 HOLD_CONFLICT και λοιπών εγκρίσεων· CI verification και native parity. Καμία νέα κλινική έγκριση. Native DEMO PENDING.

## 4.39 Ιδιωτικό R5 με machine-readable dispositions — 17/11/2026

Δημιουργήθηκε πραγματικά το ιδιωτικό `DIOMIL_REVIEW_TECHNICAL_DRAFT_20261117_R5.xlsx` με artifact_tool από R4. Προστέθηκε αποκλειστικά το φύλλο ΕΓΚΡΙΣΕΙΣ_ΧΩΡΙΣ_ΩΡΑ με ακριβείς headers του §4.38 και 37 μοναδικές source-bound αποφάσεις (8 IGNORE_OCCUPATIONAL_MARK, 29 PSYCHOLOGY_TIME_PENDING). Επαληθεύτηκε το source coordinate/αρχική τιμή κάθε εγγραφής μέσω OOXML, το ακριβές hash binding ιδιωτικά, η ακεραιότητα ZIP, η ισότητα όλων των nonempty cell values των 8 προϋπαρχόντων φύλλων και η αμεταβλητότητα SHA όλων των input XLSX. 846 weekly, 106 identities, 29 active holds διατηρούνται: δεν αφαιρούνται μέχρι να υποστηρίζεται πλήρως η διαδρομή untimed metadata και admission. Καμία κλινική έγκριση πέρα από τις ήδη ρητές αποφάσεις χειριστή, κανένα πλήρες CLI/CI/native test. R5 παραδίδεται ιδιωτικά. Επόμενο: source-bound corrections και ομαδοποίηση πραγματικών overlaps/shared slots, όχι άλλη γενική επανεξέταση. Native DEMO PENDING.

## 4.40 Ανεξάρτητη ομαδοποίηση R5 — 17/11/2026

Ο χειριστής επιβεβαίωσε παραλαβή R5. Read-only OOXML grouping όλων των 846 weekly proposals: 18 patient/day/exact-start groups, 36 συμμετέχουσες γραμμές έναντι 28 HOLD_CONFLICT flags, 10 distinct source patients. 83 provider/day/exact-start groups με πάνω από έναν ασθενή μέσω ανεξάρτητης accent/case/alphanumeric normalization, όχι canonical provider validation ή capacity proof. Καμία αυθαίρετη αφαίρεση ή μετακίνηση. Επόμενο: ιδιωτική ομαδοποιημένη λίστα αποφάσεων, διάκριση πραγματικά ταυτόσημων διπλοεγγραφών από διακριτές θεραπείες, αναθεώρηση flags με exact membership και source-bound approvals. R5 δεν τροποποιήθηκε. Native DEMO PENDING.

## 4.41 Ιδιωτικό συγκεντρωτικό φύλλο αποφάσεων — 17/11/2026

Δημιουργήθηκε το ξεχωριστό ιδιωτικό DIOMIL_ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ_20261117.xlsx από read-only OOXML ανάγνωση R5, με artifact_tool, φύλλα ΟΔΗΓΙΕΣ, ΑΣΘΕΝΕΙΣ_ΕΠΙΚΑΛΥΨΕΙΣ (18 ομάδες, δύο εγγραφές ανά ομάδα, 4 ίδιας ειδικότητας), ΚΟΙΝΑ_SLOTS (83 ομάδες, ακριβή μέλη/πηγές). Κίτρινα πεδία και dropdown αποφάσεων. ZIP integrity και αριθμός γραμμών επαληθεύτηκαν, R5 SHA unchanged. Δεν εκτελέστηκε visual render/full value-by-value readback, καμία αυτόματη έγκριση. Ομαδοποίηση κοινών slots ενδεικτική, όχι provider registry/capacity verification. Παραδόθηκε ιδιωτικά, όχι στο repo. Ζητείται πρώτα επίλυση 18 patient collisions, όχι μαζική αποδοχή των 83 ομάδων. Παραμένουν οι υπόλοιπες gates του §4.39. Native DEMO PENDING.

## 4.42 Ιδιωτικό R6 μετά τις 18 αποφάσεις επικαλύψεων — 18/11/2026

Παρελήφθη το συμπληρωμένο ιδιωτικό φύλλο με 18 απαντήσεις. Ο χειριστής επιβεβαίωσε ρητά τη δεύτερη εκδοχή των δύο αντικρουόμενων εβδομαδιαίων διορθώσεων Εργοθεραπείας. Δημιουργήθηκε με artifact_tool το DIOMIL_REVIEW_TECHNICAL_DRAFT_20261118_R6.xlsx. Ελέγχθηκε ακριβής αντιστοίχιση των 18 ομάδων με R5 (ειδικότητες, source refs, ημέρα/ώρα), αφαιρέθηκαν 4 διπλές προτάσεις και 2 επιπλέον προτάσεις με επιλογή διατήρησης A, προστέθηκε 1 ημέρα, εφαρμόστηκαν οι ρητές αλλαγές εβδομαδιαίου προγράμματος. Σύνολο 841 rows: 823 PROPOSED_EXACT_SOURCE, 8 PROPOSED_OPERATOR_CORRECTION, 8 HOLD_APPROVED_COMBINED, 1 HOLD_CONFLICT, 1 HOLD_RESCHEDULE. Πέντε exact-start groups: τέσσερις ρητά επιτρεπόμενοι ΦΘ/Ρομποτικής με ίδιο provider (διατηρούνται και οι δύο εγγραφές, εκκρεμεί κώδικας admission/capacity), μία Λογοθεραπεία με εκκρεμή συνεννόηση. Νέα φύλλα ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ (23 audit rows), ΑΡΧΙΚΕΣ_ΑΠΑΝΤΗΣΕΙΣ (18 ακριβείς απαντήσεις), ΕΝΕΡΓΕΣ_ΕΠΙΚΑΛΥΨΕΙΣ (5 groups). Παλαιό ΕΠΙΚΑΛΥΨΕΙΣ διατηρείται ως ιστορικό, όχι τρέχον gate input. Επαληθεύτηκαν ZIP integrity, πλήρης ordered weekly-cell readback και αμεταβλητότητα SHA όλων των εισόδων. Δεν έγινε visual render, πλήρες pytest/CI, native Excel ή approval ταυτοτήτων/shared slots. Το R6 παραμένει τεχνικό draft: οι ρητές διορθώσεις δεν είναι αυτόματα αποδεκτές από υπάρχον admission και reconciliation. Επόμενο: source-bound correction ledger, ακριβής εξαίρεση συνδυασμού ΦΘ/Ρομποτικής χωρίς γενική άδεια επικάλυψης ή διπλή κατανάλωση capacity, επανυπολογισμός κοινών slots και έλεγχος tests. Καμία ιδιωτική τιμή στο repo. Native DEMO PENDING.

## 4.43 Ρητή εξαίρεση σε εβδομαδιαίο συνδυασμό ΦΘ/Ρομποτικής — 18/11/2026

Ο χειριστής επιβεβαίωσε ότι πρόκειται για εξαίρεση και υπενθύμισε τον ήδη υπάρχοντα μηχανισμό επιβεβαίωσης. Εξετάστηκαν double_booking_exception.py, apply_confirmed_double_booking.py, prepare_double_booking_review_ui.py, patient_schedule.py και trial_provider_capacity.py. Ο operational μηχανισμός ημερήσιας αντικατάστασης παραμένει αμετάβλητος: χαλαρώνει μόνο therapist occupancy, όχι patient collision. Ο υπάρχων weekly validator επεκτάθηκε στο commit 1cef01f με προαιρετικό confirmed_combined_slots exact frozenset pairs, default empty/strict. Μόνο δύο διαφορετικές εγγραφές ΦΘ/Ρομποτικό, ίδιος ασθενής/weekday/time, ίδιος μη κενός normalized provider και ακριβές confirmed pair επιτρέπονται. Stale/third/duplicate/wrong-specialty αποφάσεις απορρίπτονται. Commit 0048929 πρόσθεσε δύο synthetic tests με θετικές/αρνητικές περιπτώσεις. Δεν εκτελέστηκαν ακόμη. Ο υπάρχων capacity audit ήδη μετρά distinct provider/day/clock. **Δεν έχει συνδεθεί ακόμα** το optional confirmation με admission, plan, persisted weekly reader, native UI ή source-bound R6 ledger. Δεν υπάρχει νέα γενική άδεια ούτε ενεργοποίηση των 4 R6 συνδυασμών. Επόμενο: persistence exact approval μέσω existing confirmation pattern και propagation με fail-closed readback, επαλήθευση tests και daily/native parity. R6/operational αρχεία αμετάβλητα. Native DEMO PENDING.


## 4.44 Νέα δεσμευτική λειτουργική απαίτηση: παράλληλος προγραμματισμός σημερινής και μελλοντικών ημερών — 19/11/2026

Ο χειριστής περιγράφει τη φυσιολογική καθημερινότητα του κέντρου: πριν ολοκληρωθούν οι σημερινές θεραπείες, οργανώνεται το αύριο, με ήδη γνωστές άδειες θεραπευτών, προσωρινά εξιτήρια/απουσίες, εξετάσεις ασθενών και προγραμματισμένες αλλαγές. Η τρέχουσα ημέρα συνεχίζει να δέχεται έκτακτες διακομιδές, αλλαγές ώρας/θεραπευτή κ.λπ. Απαίτηση: σχεδιασμός, επιβεβαίωση και durable αποθήκευση μελλοντικής ημέρας, επάνοδος στη σημερινή χωρίς απώλεια, ανάκτηση του προετοιμασμένου προγράμματος όταν έρθει η ημέρα και διαχείριση μόνο των νέων έκτακτων αλλαγών. Αφορά πλήρη ωριαία προβολή όλων των ειδικοτήτων και θεραπευτών, διαχείριση προγραμματισμού ειδικά ΦΘ και συναφών κατά τους ήδη συμφωνημένους κανόνες.

### Υφιστάμενη βάση και πραγματικό κενό
Read-only source inspection του branch: models.py (Session.session_date, DailyAbsence.absence_date/covers, DailySessionCancellation.cancellation_date, ReplacementAssignment.replacement_date), daily_input_reader.py (ημερομηνία DAILY_INPUT!B2 και date-keyed changes), daily_input_sheet.py / preview_daily_input_sheet.py (target_date), build_finalized_day_snapshot_preview.py (service_date από B2), install_unified_day_save_ui.py (Save Day επιλέγει CloseOperationalDay ή StartClosedDayAmendment βάσει ημερομηνίας). CHECKPOINT_H6_MASTER_ONLY_SAVE_2026-11-03.md καταγράφει native proof same-day v1/v2 και MASTER-only save. Ο κεντρικός οδηγός §§2–4 περιγράφει recurring baseline, date-effective lifecycle και χωριστά ημερήσια overlays. Άρα υπάρχει date-aware αρχιτεκτονική, αλλά δεν αποδείχθηκε ακόμη πλήρης native εναλλαγή πολλών ανοικτών ημερών, διατήρηση όλων των draft/confirmed overlays ή σωστή επαναφόρτωση μετά από save/reopen. Το B2 μόνο δεν είναι έγκριση ασφαλούς date switching. Ιδίως το υπάρχον Save Day είναι finalization/amendment, όχι future draft save. Δεν αλλάζουμε παλαιά frozen proofs H2–H10.

### Προτεινόμενη UX (όχι ακόμη υλοποιημένα controls)
Στο ίδιο authoritative XLSM/MASTER, καθαρή ένδειξη επιλεγμένης ημερομηνίας και κατάστασης, Σήμερα / Αύριο / Άλλη ημέρα, χωρίς δεύτερο ανεξάρτητο workbook ή ανάγκη χειροκίνητης αλλαγής B2. Να εξεταστεί next-working-day shortcut για ΔΕ–ΠΑ και διακριτή literal calendar επιλογή. Όλες οι σχετικές προβολές και operations να ακολουθούν την επιλεγμένη service date. Αποθήκευση προετοιμασίας/προγράμματος χωριστά από Οριστικοποίηση ημέρας. Οι ακριβείς labels/placement θα συμφωνηθούν με βάση υπάρχον layout. Όχι αυτόματη καταγραφή attendance/history για μελλοντικό πρόγραμμα, ούτε υποχρεωτικό κλείσιμο σημερινής για πρόσβαση αύριο. Επιλεγμένη service date και πραγματική ημερομηνία καταγραφής δεν συγχέονται.

### Acceptance πριν το πρώτο λειτουργικό DEMO
Επαναχρησιμοποίηση υφιστάμενων operational actions/ledgers, όχι δεύτερη μηχανή ή νέα ανεξάρτητη πηγή. Προτεραιότητα στην επαλήθευση των date keys/persistence/readback και των action routes. Σε isolated synthetic native XLSM: (1) δύο ανοικτές ημερομηνίες, διαφορετικές απουσίες/αλλαγές/εγκεκριμένες αντικαταστάσεις, (2) Σήμερα→Αύριο→Σήμερα με ακριβή επαναφορά, χωρίς δεδομένα της άλλης ημέρας ή αλλαγή recurring baseline, (3) έκτακτη σημερινή αλλαγή μετά την έγκριση αυριανού χωρίς απώλεια αύριο, (4) save/reopen με τα δύο overlays και τις αποφάσεις ακριβώς intact, (5) οριστικοποίηση μόνο της σημερινής, χωρίς future attendance/finalization, (6) επόμενη πραγματική εργάσιμη ημέρα: ανάκτηση ήδη προετοιμασμένου προγράμματος, νέες έκτακτες αλλαγές, χωρίς stale assignments· επανέλεγχος επικαιροποιημένων date-effective census/recurring events και availability με ευδιάκριτη επισήμανση συγκρούσεων, όχι σιωπηλή εξαφάνιση προηγούμενων αποφάσεων ή αυτόματη γενίκευση εξαιρέσεων, (7) όλες οι σχετικές προβολές ημερομηνιακά συνεπείς, περιλαμβανομένων MASTER/DAILY_INPUT/THERAPIST_DAILY. Να διατηρούνται οι υφιστάμενες σημασιολογίες προσωρινών απουσιών vs πραγματικού εξιτηρίου· καμία αυθαίρετη κλινική υπόθεση. Όχι finalization μελλοντικής ημέρας μέσω συνήθους draft flow. Πρόκειται για υποχρεωτικό λειτουργικό acceptance criterion, όχι νέο αίτημα επανάληψης των frozen tests. Καμία αλλαγή κώδικα/workbook ούτε native δοκιμή σε αυτό το checkpoint. R6 και §§4.42–4.43 παραμένουν οι εκκρεμότητες εισαγωγής· νέο distinct date-navigation/persistence gate καταγράφεται για DEMO.

## 4.45 Πρώτη υλοποίηση και εκτελεσμένα focused tests παράλληλων ημερών — 19/11/2026

Ο χειριστής ζητά να αρχίσει αυτόνομα η εργασία/στοχευμένα tests όσο βρίσκεται μακριά από Windows PC. Στο ενεργό branch εξετάστηκαν prepare_daily_input_in_place.py, daily_input_reader.py, daily_input_sheet.py, daily_replacement_facts.py, daily_state.py, day_finalization.py, install_unified_day_save_ui.py, prepare_daily_change_restore_in_place.py, patch_native_restore_recovery.py και το ιστορικό RECOVERY_DAILY_CHANGE_RESTORE_2026-10-31.md. Επιβεβαιώθηκε: DAILY_INPUT!B2 μία ενεργή ημερομηνία, editable θεραπευτές A:F και ασθενείς A:E με δυναμικά row bounds, ημερομηνιακά change IDs, αντικαταστάσεις ήδη persistent date-aware, generalized ledger περιέχει view/audit fields και δεν αρκεί για lossless reconstruction ολόκληρων raw input rows (π.χ. χωριστό σχόλιο/από-έως/ημιτελή στοιχεία). Το υπάρχον prepare_daily_input_payload δεν επιτρέπει clean day χωρίς ειδική διαδρομή· δεν το χαλαρώνουμε αυθαίρετα. Η ιστορική restore patch έχει unscoped display iteration· χρειάζεται εξακρίβωση του canonical native module/schema πριν στοχευμένη ημερομηνιακή UI/restore αλλαγή.

Πραγματικά commits:
- 15e45d9: daily_state cancellation/replacement lookup με (session ID, service date), όχι μόνο ID. Υφιστάμενη προτεραιότητα cancellation/patient absence/replacement/therapist absence αμετάβλητη.
- 1418109: core build_finalized_day_snapshot απορρίπτει μελλοντικό service date ως προς την ημερομηνία του timezone-aware finalized_at. c166745: MASTER SaveOperationalDay dispatcher προειδοποιεί/σταματά για serviceDate > local VBA Date πριν close/amendment. Native installer/Excel δεν εκτελέστηκε.
- f6b1cfb: δύο synthetic regressions στο υπάρχον test_day_finalization.py: δύο ημερομηνίες, distinct confirmed replacements/cancellations, μετάβαση/late absence, σήμερα-only finalization, future rejection και αποδοχή την αντίστοιχη ημερομηνία.
- c2fc5d8: Python/rehab_excel/daily_input_day_switch.py pure open-day switch plan με immutable raw DAILY_INPUT grids (διατήρηση τύπων, ακριβών row positions, κενών/ημιτελών input, dynamic layout, expected active date, source/target not finalized). Current grid ενημερώνει μόνο τη δική του ημερομηνία, target grid ανακτάται ή αρχικοποιείται blank, other-date drafts αμετάβλητα. Δεν είναι δεύτερο scheduling/attendance/exception authority. No native storage or writer; caller πρέπει να επαληθεύει formulas/errors, exact history status/layout/source hash και να κάνει atomic transaction/readback. Δεν επιτρέπει ακόμη transition από/προς finalized day μέσω αυτού του open-only path.
- 9c2d9b9: Tests/test_daily_input_day_switch.py 13 synthetic cases, roundtrip με τρέχουσες/μελλοντικές αλλαγές, τρίτη ημέρα, stale/duplicate/layout/finalized/timestamp/invalid values.

**Πραγματική εκτέλεση:** direct git fetch στο test container απέτυχε DNS. Ανασυστάθηκαν προσωρινά τα συγκεκριμένα οκτώ Python source/test files από το connector, ελέγχθηκαν με git hash-object και ταυτίστηκαν ακριβώς με τα GitHub blob SHAs: models 91306a4, daily_state 5a3cfbc, daily_statistics c6878b06, day_finalization 2a7b29b, orphan_resolution 73b838c, test_day_finalization 54a35e4, daily_input_day_switch b6ea425, test_daily_input_day_switch 7b47ed2. Minimal empty package initializers στο isolated runtime, όχι πλήρες checkout. `python -m pytest test_day_finalization.py test_daily_input_day_switch.py -q`: **19 passed in 0.06s**, 6+13, source-only, όχι full suite/CI/native. Workflow runs/combined status για 9c2d9b9 επέστρεψαν κενά· no CI GREEN claim. Όλα fictional, κανένα real workbook τροποποιημένο.

**Επόμενο ακριβές engineering:** read-only capture/validation literal editable zones από actual DAILY_INPUT spec, immutable source-date archive μέσα στο ίδιο workbook (να εξεταστεί reuse υπάρχοντος ledger χωρίς loss semantics, όχι νέο ανεξάρτητο XLSM), verified exact-date archive load και transactional native apply/rollback. Συνδυασμός με υπάρχον date-keyed ACTIVE replacements/restore/ledger και projection rebuild για clean/changed dates, ημερομηνιακή συνέπεια MASTER/DAILY_INPUT/THERAPIST_DAILY, επιλογή/μετάβαση από finalized προς open day χωρίς ιστορική μετάλλαξη. Προτού χαρακτηριστεί native accepted, full focused regressions/CI, νέα isolated Windows synthetic workbook και end-to-end proof §4.44, όχι επανάληψη frozen tests. Ο χρήστης δεν απαιτείται να εκτελέσει κάτι τώρα. R6 admission/combined exception pending παραμένουν χωριστά. Native DEMO PENDING.


## 4.46 Διαγράμματα ροής παράλληλου ημερήσιου προγραμματισμού — 19/11/2026

Ο χειριστής ζήτησε τα αντίστοιχα πρακτικά flowcharts για τις §§4.44–4.45. Στον κεντρικό οδηγό, PROJECT_DELIVERY_MASTER_REPORT_20261115 και CURRENT_HANDOFF δεν βρέθηκαν ειδικά flowcharts για αυτό το νέο σενάριο. Τα ακόλουθα είναι **προτεινόμενα, δεσμευμένα με τα υπάρχοντα contracts διαγράμματα υλοποίησης/acceptance, όχι απόδειξη ολοκληρωμένης native λειτουργίας**. Η επιλογή ακριβούς μορφής controls, διαλόγων και storage schema μένει προς υλοποίηση/επαλήθευση. Μία authoritative εφαρμογή XLSM, όχι διαφορετικά workbook ανά ημέρα.

### A. Αλλαγή επιλεγμένης ημερομηνίας (MASTER)
```mermaid
flowchart TD
    A["MASTER: ενεργή ημερομηνία και κατάσταση"] --> B["Σήμερα / Επόμενη εργάσιμη / Άλλη ημερομηνία"]
    B --> C{"Διαφορετική ημερομηνία;"}
    C -->|Όχι| A
    C -->|Ναι| D{"Εκκρεμούν μη αποθηκευμένες αλλαγές;"}
    D -->|Ναι| E{"Εφαρμογή ή αποθήκευση πρόχειρου ή ακύρωση μετάβασης"}
    E -->|Ακύρωση| A
    E -->|Εφαρμογή ή πρόχειρο| F["Έλεγχος τρέχουσας ημερομηνίας, layout, source hash και κατάστασης"]
    D -->|Όχι| F
    F --> G{"Στόχος ήδη οριστικοποιημένος;"}
    G -->|Ναι| H["Ιστορική προβολή χωρίς εγγραφή / ξεχωριστή amendment route"]
    G -->|Όχι| I["Ανάκτηση ημερήσιου draft και τελευταίων επιβεβαιωμένων facts ή καθαρού overlay"]
    I --> J["Επανυπολογισμός με ισχύον recurring, census, roster, absences, replacements"]
    J --> K{"Νέες συγκρούσεις ή stale επιβεβαιώσεις;"}
    K -->|Ναι| L["Εμφάνιση εκκρεμοτήτων, καμία σιωπηλή διαγραφή ή έγκριση"]
    K -->|Όχι| M["Προετοιμασία date-specific προβολών"]
    L --> M
    M --> N["Atomic native apply / semantic readback / save"]
    N --> O{"Επιτυχία;"}
    O -->|Ναι| P["MASTER + DAILY_INPUT + THERAPIST_DAILY ίδια ημερομηνία"]
    O -->|Όχι| Q["Rollback στην προγενέστερη κατάσταση, διατήρηση drafts και σαφές μήνυμα"]
```

Ο μηχανισμός open-only του §4.45 καλύπτει μόνο το pure plan για δύο μη οριστικοποιημένες ημερομηνίες. Η ιστορική προβολή/μετάβαση από κλειστή σε ανοικτή απαιτεί ξεχωριστό read-only/transaction route. Μη αποθηκευμένο ή ημιτελές input δεν εξισώνεται με εγκεκριμένη λειτουργική αλλαγή: πρέπει να διατηρηθούν τα raw editable δεδομένα και η τελευταία επιβεβαιωμένη effective κατάσταση, χωρίς αυτόματη ενεργοποίηση μη έγκυρης γραμμής. Αν αποτύχει validation ή persistence, δεν αλλάζει η ενεργή ημερομηνία. Καμία αυτόματη finalization στην εναλλαγή. Weekend shortcut: επόμενη εργάσιμη ΔΕ–ΠΑ, το ακριβές label/holiday behaviour μένει προς επιβεβαίωση.

### B. Προετοιμασία μελλοντικής ημέρας και εφαρμογή
```mermaid
flowchart TD
    A["Επιλογή μελλοντικής εργάσιμης ημερομηνίας"] --> B["Φόρτωση ισχύοντος εβδομαδιαίου προγράμματος και αποθηκευμένου overlay"]
    B --> C["DAILY_INPUT: γνωστές άδειες / εξετάσεις / προσωρινές απουσίες"]
    C --> D{"Έγκυρες καταχωρίσεις;"}
    D -->|Όχι| E["Διόρθωση ή αποθήκευση μη εφαρμοσμένου πρόχειρου"]
    E --> C
    D -->|Ναι| F["Υπάρχων μηχανισμός ημερήσιας προβολής και προτάσεων ΦΘ"]
    F --> G{"Χρειάζεται εξαίρεση ή παραμένει εκκρεμότητα;"}
    G -->|Ναι| H["Ρητή επιβεβαίωση μόνο επιτρεπτής εξαίρεσης ή διατήρηση εκκρεμότητας"]
    G -->|Όχι| I["Προεπισκόπηση και επιβεβαίωση επιλογών"]
    H --> I
    I --> J["Αποθήκευση προγράμματος ανά service date και verified readback"]
    J --> K["Επιστροφή στο Σήμερα χωρίς αλλαγή σημερινών facts"]
```

Χρήση υπαρχόντων solver, capacity, recurring/patient conflict, double-booking confirmation και replacement facts με τις ήδη ισχύουσες δικλίδες, όχι νέος ανεξάρτητος scheduler. Η συνήθης αποθήκευση μελλοντικού προγράμματος δεν καταγράφει attendance, CURRENT_FINAL ή DAY_FINALIZED. Εκκρεμής μη επιλυμένη συνεδρία παραμένει εμφανής, όχι αυτόματα ακυρωμένη. Μελλοντικό μη παρατηρημένο no-show δεν δηλώνεται ως πραγματοποιημένο γεγονός. Οι μόνιμες αλλαγές και πραγματικά εξιτήρια έχουν χωριστή date-effective σημασιολογία.

### C. Έναρξη νέας πραγματικής ημέρας και οριστικοποίηση
```mermaid
flowchart TD
    A["Άνοιγμα DIOMIL ή επιλογή Σήμερα"] --> B["Πραγματική τοπική ημερομηνία"]
    B --> C{"Υπάρχει οριστικοποιημένο ιστορικό;"}
    C -->|Ναι| D["Ιστορική προβολή / ρητή amendment route"]
    C -->|Όχι| E["Ανάκτηση αποθηκευμένου προγράμματος ή αρχικοποίηση από recurring"]
    E --> F["Έλεγχος επίκαιρου census, αναθέσεων, διαθεσιμότητας και εξαιρέσεων"]
    F --> G["Διατήρηση προηγούμενων έγκυρων επιλογών και επισήμανση νέων συγκρούσεων"]
    G --> H["Έκτακτες αλλαγές / επιβεβαιώσεις με τον υπάρχοντα ημερήσιο μηχανισμό"]
    H --> I{"Ο χρήστης ζητά οριστικοποίηση;"}
    I -->|Όχι| J["Αποθήκευση προγράμματος / συνέχεια εργασίας"]
    J --> H
    I -->|Ναι| K{"Σωστή ημερομηνία και πλήρη έγκυρα αποτελέσματα;"}
    K -->|Όχι| H
    K -->|Ναι| L["Υπάρχων CloseOperationalDay / verified history"]
    L --> M["CURRENT_FINAL μόνο για τη συγκεκριμένη service date"]
```

Η πραγματική ημερομηνία ανοίγματος δεν ταυτίζεται με την επιλεγμένη ημερομηνία επεξεργασίας. Εκκρεμής προηγούμενη ημέρα δεν κλείνει αυτόματα κατά την αλλαγή ημέρας. Όλα τα μελλοντικά drafts και confirmed overlays παραμένουν ανεπηρέαστα. Μελλοντικό service date απορρίπτεται από τον ήδη προστεθέντα source guard (§4.45); επιβεβαίωση ολοκλήρωσης πραγματικής ημέρας εξακολουθεί να απαιτείται.

### D. Τεχνικό state/ownership και κρίσιμη συναλλαγή

- Recurring/registry/date-effective events: υφιστάμενη authoritative πηγή.
- Raw DAILY_INPUT ανά service date: αποθήκευση/ανάκτηση χωρίς απώλεια literal πεδίων και ακριβούς row identity. Δεν αντικαθιστά confirmed facts.
- Confirmed ημερήσια facts/approved exceptions: υφιστάμενα date-keyed ledgers ή στοχευμένη lossless επέκτασή τους. Καμία διπλή πηγή αλήθειας ή γενική έγκριση.
- Operational προβολές: παράγονται ξανά από την ενεργή ημερομηνία και την effective κατάσταση. Ποτέ authoritative αντίγραφο ωρών από rendered cells.
- Historical facts: immutable finalized snapshots και υπάρχουσα versioned amendment route.

Η αλλαγή ημερομηνίας απαιτεί atomic validation → prepare → apply → persisted readback → commit, αλλιώς rollback χωρίς απώλεια του outgoing draft. Restore row/selected/all περιορίζεται στην επιλεγμένη ημερομηνία και στα ακριβή change IDs. Απαιτούνται δοκιμές εναλλαγής clean/dirty, failure injection, save/reopen, σήμερα-only finalization, history read-only, επόμενη εργάσιμη, late αλλαγές σε census/provider, stale approvals, προστασία unresolved drafts και συμφωνία ημερομηνιών σε όλα τα operational views. Δεν πρέπει να δεσμευτεί πραγματικό slot από πρόχειρη μη επιβεβαιωμένη καταχώριση.

**Κατάσταση:** η παρούσα ενότητα προσθέτει μόνο τεκμηρίωση των ροών, όχι νέο code/test/native proof. Παραμένουν οι 19 πραγματικά εκτελεσμένοι focused source-only έλεγχοι του §4.45. Επόμενο engineering: read-only capture/store/readback raw DAILY_INPUT zones και διαχωρισμός draft/confirmed, με επαναχρησιμοποίηση των existing ledgers, πριν date navigation/UI. R6 import ξεχωριστά pending. Native DEMO PENDING.

## 4.47 Πρακτικές διαδρομές χειριστή και contextual επιλογές — 19/11/2026

Ο χειριστής διευκρίνισε ότι δεν αρκούν τα τεχνικά flowcharts: κάθε λειτουργία πρέπει να είναι ευδιάκριτη και προσβάσιμη από τη φυσική διαδρομή εργασίας, με τις σωστές επιλογές τη στιγμή της ανάγκης. Έλεγχος source: master_toolbar_vba.py έχει ήδη Απουσίες/Ακυρώσεις → DAILY_INPUT, Ημερήσιο πρόγραμμα → THERAPIST_DAILY, Μητρώο → ShowRegistrationMenu, BackToMaster και Save & Exit. daily_input_action_vba.py::ApplyDailyInputPreview ήδη εκτελεί validation, apply, refresh, ResolveDailyReplacementsInPlace όταν χρειάζεται, και GoToMaster στην επιτυχία. install_unified_day_save_ui.py::SaveOperationalDay είναι close/amendment dispatcher με νέο future-date guard, όχι draft save. Τα υπάρχοντα controls και παλαιά native proofs δεν ανακατασκευάζονται αυθαίρετα.

Προτεινόμενο operator routing contract (labels/θέσεις προς native UI επαλήθευση):
- MASTER: ευδιάκριτη επιλεγμένη service date και κατάσταση, Σήμερα / Αύριο ή Επόμενη εργάσιμη / ημερολόγιο. Κοινή ημερομηνία σε όλες τις operational οθόνες, επιστροφή MASTER χωρίς ακούσια αλλαγή της.
- Σήμερα ή μελλοντική ημέρα → Απουσίες/Ακυρώσεις → επιλογές θεραπευτή/ασθενούς, όλης ημέρας ή ώρας/διαστήματος σύμφωνα με υπάρχον input contract → Εφαρμογή & Συνέχεια → προτάσεις μόνο όταν απαιτούνται → επιλογή/επιβεβαίωση επιτρεπτής λύσης ή εκκρεμότητα → MASTER ίδιας ημερομηνίας. Έκτακτη διακομιδή ως προσωρινή απουσία μόνο όταν πράγματι προσωρινή, οριστικό εξιτήριο χωριστά.
- Σύγκρουση: επιτρεπτή εξαίρεση εμφανίζει συγκεκριμένη προειδοποίηση και ρητή επιβεβαίωση ή επιστροφή για άλλη λύση. Hard constraints δεν παρακάμπτονται. Χειροκίνητη αλλαγή ώρας/θεραπευτή να είναι προσβάσιμη από το σχετικό ημερήσιο context ακόμη χωρίς αρχική απουσία, με τους υπάρχοντες κανόνες· χρειάζεται επιβεβαίωση end-to-end δυνατότητας, όχι ισχυρισμός ότι υπάρχει ήδη για όλες τις θεραπείες.
- Επαναφορά συγκεκριμένης/επιλεγμένων/όλων των ημερήσιων αλλαγών μόνο της επιλεγμένης ημερομηνίας. Οι ήδη γνωστοί περιορισμοί restore ανά τύπο παραμένουν μέχρι verified υποστήριξη.
- Μητρώο → επεξεργασία μόνιμου προγράμματος/effective date για recurring αλλαγές και lifecycle, όχι ημερήσιο overlay. Αναφορές/ιστορικό προσβάσιμα από existing navigation χωρίς μετατροπή history σε editable draft.
- Αποθήκευση προγράμματος: durable draft και confirmed facts χωρίς history finalization. Οριστικοποίηση ημέρας: χωριστή, σαφής ενέργεια μέσω υπάρχοντος dispatcher, όχι για μελλοντική ημέρα. Σε ήδη οριστικοποιημένη ημερομηνία read-only history και σαφής, ξεχωριστή διορθωτική διαδρομή. Save & Exit διατηρεί την εργασία, δεν οριστικοποιεί από μόνο του. Οι ακριβείς νέες ετικέτες δεν έχουν οριστικοποιηθεί.
- Επιλογές εμφανίζονται όταν σχετίζονται με την κατάσταση εργασίας: όχι όλες ως κουμπιά στο MASTER, όχι τεχνικά IDs, κρυφά φύλλα, scripts ή χειροκίνητη πλοήγηση Excel. Pending draft, ανεφάρμοστη αλλαγή, επιβεβαιωμένο πρόγραμμα και ιστορικά οριστικό αποτέλεσμα έχουν διακριτή σημασιολογία. Αποτυχία οδηγεί σε κατανοητό μήνυμα και ασφαλή επιστροφή/rollback.

Acceptance: operator walk-through από την κεντρική οθόνη χωρίς terminal ή αλλαγή κρυφών κελιών, για τα σενάρια §§4.44–4.46, έλεγχο κάθε control binding/state/visibility/return route, σαφή ημερομηνία, αποθήκευση/ακύρωση/εξαίρεση/restore, late changes και save/reopen. Επιβεβαίωση πραγματικής λειτουργίας, όχι μόνο screenshots/παρουσία controls. Πρόκειται για UX specification, όχι νέα υλοποίηση ή native proof. Δεν τροποποιήθηκαν κώδικας ή workbook. Επόμενο engineering παραμένει §4.45–4.46 raw capture/archive/readback, date-bound native transaction και προβολές πριν εγκατάσταση των νέων controls. Native DEMO PENDING.

## 4.48 Literal DAILY_INPUT archive και source-bound date-switch preparation — 19/11/2026

Συνέχεια §§4.44–4.47 χωρίς νέα operator έγκριση. Επιθεωρήθηκαν daily_input_sheet.py (named-list validations PY_THERAPISTS/PY_PATIENTS, editable A:F/A:E, variable layout), daily_input_reader.py, application_contract.py (DAILY_CHANGE_LEDGER schema), finalized_day_snapshots.py και τα υπάρχοντα pure day-switch tests. Το existing ledger δεν διατηρεί χωριστά όλα τα literal form fields, σχόλια, incomplete rows ή αρχικούς τύπους· δεν είναι lossless draft store. Επιλέχθηκε νέο πολύ κρυφό, προστατευμένο βοηθητικό sheet DAILY_INPUT_DRAFTS μέσα στο ίδιο XLSM για raw forms μόνο, όχι δεύτερος scheduler, confirmed facts, capacity ή history authority. Η native δημιουργία του δεν έχει ακόμη εκτελεστεί. Το encoding δεν είναι encryption ή ανωνυμοποίηση.

Commits:
- f85d237, f7f9646, 8f8f5d2: daily_input_draft_archive.py. Capture μέσω πραγματικών named-list range boundaries, expected header/layout, canonical date parser callback, literal typed grid, partial fields και blanks. Reject formulas/errors, ambiguous ranges, unexpected populated rows, excessive bounds. Archive rows ServiceDate/Part/Parts/PayloadSHA256/Payload, typed JSON με date binding, base64 prefixed text chunks κάτω από Excel cell limit, complete chunk/digest checks, duplicate/wrong-date corruption rejection. Digest αφορά payload integrity, όχι authentication ή ανίχνευση διαγραφής ολόκληρου date record χωρίς expected membership. Native writer οφείλει να κάνει literal rehydration (όχι formula interpretation) και πλήρη independent expected-row readback.
- aac95af, 0bac49a: Tests/test_daily_input_draft_archive.py. Εκτελέστηκαν πράγματι **17 passed in 0.17s** σε isolated local Python runtime με ανακατασκευασμένα αρχεία και επιβεβαιωμένα git blob SHAs: daily_input_day_switch b6ea4253b3fb01f7e55f09bdded25f493be01524, daily_input_draft_archive 41d9016168e0b60053b91e44a427f40cef27afd2, test_daily_input_draft_archive 88441d116fcbcd87ea77fe57f576a1ce5a23caee. Συμπεριλαμβάνεται in-memory synthetic XLSX serialization/reopen του protected veryHidden archive, exact typed readback, formula/error rejection και corrupt chunk/date tests. Όχι native XLSM/COM, πλήρες checkout ή CI. Τα 19 προηγούμενα focused tests §4.45 δεν επαναλήφθηκαν.
- 4bb28b1: daily_input_switch_preparation.py, read-only XLSM input, existing _parse_date/read_finalized_day_snapshots/file_sha256, explicit first-install flag για absent archive (ποτέ ως recovery), protected veryHidden existing archive, expected selected date, source hash before/after, pure plan και expected archive rows. Όλες οι dates με ιστορικά snapshot rows, ακόμη superseded-only, απαγορεύονται στην open-only route. Δεν εφαρμόζει, δεν οριστικοποιεί, δεν γράφει αρχείο. Private result fields repr=False, όχι logging. 8c015e4: 7 πρόσθετα source-bound integration tests με synthetic VBA-free workbook fixtures, δεν εκτελέστηκαν στο παρόν runtime. Workflow runs/combined status για 8c015e4 κενά· CI unverified.

**Επόμενο ακριβές engineering:** verified native transaction (preflight source SHA/date/layout/expected full archive membership, protected/veryHidden storage provisioning only on explicit first install, literal row writes, persisted readback και rollback), διακριτή διατήρηση last-applied confirmed state vs raw draft, date-filtered change ledger/restore and reconstruction via existing scheduling/projection machinery. Κανένα UI date toggle πριν αποδειχθεί ότι η μετάβαση δεν αφήνει MASTER/THERAPIST_DAILY stale, δεν ενεργοποιεί incomplete draft, και δεν μεταβάλλει άλλη ημερομηνία. Να εκτελεστούν τα 7 integration tests και τα scoped regressions/CI, μετά native synthetic save/reopen και acceptance §§4.44–4.47. Απουσία native Excel/COM και πλήρους checkout στο παρόν runtime, όχι ανάγκη operator παρέμβασης ακόμη. Δεν τροποποιήθηκαν operational ή private πραγματικά αρχεία, R6 admission ανεξάρτητα pending. Native DEMO PENDING.

## 4.49 Κοινή ερμηνεία ημερήσιας φόρμας και διάκριση draft / εφαρμοσμένης κατάστασης — 19/11/2026

Συνέχεια §§4.44–4.48. Ο προγενέστερος κύκλος των 7 synthetic integration tests αναφέρθηκε επιτυχής σε isolated runtime με αποσπάσματα των dependencies· δεν ήταν πλήρες exact-source checkout ή native. Τα 7 tests παραμένουν προς πλήρη ανεξάρτητη CI επαλήθευση. Η canonical worksheet parsing λογική εξήχθη στο daily_input_reader.py χωρίς αλλαγή κανόνων (ccfbbb4), και επαναχρησιμοποιείται από interpret_daily_input_grid με ίδιο date, row geometry, validation και source change IDs (40dc16e). Test commit 20cc4df: parity ορατής/αρχειοθετημένης φόρμας, δεύτερη ημερομηνία, ελλιπές draft, clean/restore options.

Νέα commits 7b8a151 και d257330: daily_input_applied_state.py και Tests/test_daily_input_applied_state.py. Η verify_day_input απαιτεί ρητή τελευταία εφαρμοσμένη literal grid της ίδιας ημερομηνίας/layout, δεν ερμηνεύει το τρέχον ημιτελές draft ως εφαρμοσμένο, διασταυρώνει τα source changes με ACTIVE date-scoped DAILY_CHANGE_LEDGER (ChangeID, ChangeType, SubjectID, Subject, Reason, SourceRow) και διατηρεί βεβαιωμένη dirty ένδειξη με typed serialization. RESTORED και άλλες ημερομηνίες δεν αποτελούν ενεργές αλλαγές αυτής της ημέρας. Missing applied evidence, mismatch, duplicate, unsupported status/type, unsafe cells ή schema αποτυγχάνουν κλειστά. Το ledger δεν περιέχει όλα τα raw fields, άρα αυτή η αντιστοίχιση δεν πιστοποιεί αυτάρκως όλο το payload/χρονικό διάστημα και δεν είναι επιτρεπτό να ανακατασκευαστεί applied grid από το ledger ή από πρόχειρη φόρμα. Το applied evidence πρέπει να προέρχεται από επιτυχή verified native Apply/Restore και δεν υποκαθιστά το existing date-keyed replacement ledger. Καμία νέα κλινική έγκριση, workbook write ή ενεργοποίηση UI.

Πραγματική εκτέλεση σε isolated temporary Python runtime με empty package initializers και ανακατασκευασμένα ακριβή GitHub αρχεία, git hash-object ταυτίστηκε: daily_input_applied_state 256ff2651bee39e572ddd50de57879b0fc7d365c, test_daily_input_applied_state 6b4baa7b844b27c00c3edf0c78b25118cd7dc65a, application_contract 38d43ab067e43ee36836979a6010f6ffd525b0e5, daily_input_reader 19ca445de14c42751c6db674586e5afad1187c54, daily_input_day_switch b6ea4253b3fb01f7e55f09bdded25f493be01524, daily_input_draft_archive 41d9016168e0b60053b91e44a427f40cef27afd2, models 91306a46637eddfde178bf64ad43931b4666f699, test_daily_input_grid_parity e2140a385d76ab07b3e68191ac9f99cea269dacc. Command: python -m pytest test_daily_input_applied_state.py test_daily_input_grid_parity.py -q → **16 passed in 0.20s (12 νέα + 4 parity)**. PytestUnraisableExceptionWarning από ZipFile.__del__ / closed file σε συνθετικό Excel fixture, δεν απέτυχαν τα tests· όχι πλήρες checkout, remote CI ή native Excel, καμία πραγματική εγγραφή ασθενούς.

Επόμενο: η source-only verify_day_input δεν χρησιμοποιείται ακόμη από operational Apply/Restore/Refresh/CloseDay. Πρόσδεση applied evidence σε επιτυχή native συναλλαγή (με προστατευμένη ημερομηνιακή αποθήκευση, ελεγμένη αρχικοποίηση καθαρής νέας ημέρας και fail-closed για παλιές μη αποδείξιμες εγγραφές), source/target readback, διατήρηση applied state σε draft-only save, αποτροπή stale/dirty finalization και χρήση αποκλειστικά verified applied facts για rebuild προβολών. Έλεγχος όλες οι entrypoints και rollback πριν συνδεθούν τα κουμπιά. Native Excel και full regression παραμένουν pending. Τα R6 admission/exception gates παραμένουν ανεξάρτητα. Δεν απαιτείται ακόμη χειροκίνητη παρέμβαση.
