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
