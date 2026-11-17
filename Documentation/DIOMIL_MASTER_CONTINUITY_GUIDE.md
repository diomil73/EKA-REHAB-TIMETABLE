# DIOMIL | ΕΝΙΑΙΟΣ ΟΔΗΓΟΣ ΑΚΡΙΒΟΥΣ ΣΥΝΕΧΕΙΑΣ

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
