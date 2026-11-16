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
