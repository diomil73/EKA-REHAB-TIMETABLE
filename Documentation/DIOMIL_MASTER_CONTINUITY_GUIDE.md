# DIOMIL | ΕΝΙΑΙΟΣ ΟΔΗΓΟΣ ΑΚΡΙΒΟΥΣ ΣΥΝΕΧΕΙΑΣ

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
