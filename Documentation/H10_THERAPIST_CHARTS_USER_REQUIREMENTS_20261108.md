# H10 / επόμενη έκδοση: απαιτήσεις γραφημάτων θεραπευτών (καταγραφή συνομιλίας)

**Ημερομηνία καταγραφής:** 2026-11-08  
**Προέλευση:** αποφάσεις και παρατηρήσεις του χρήστη κατά τις δοκιμές BO → BR → BS και τη συζήτηση για BT / ενδεχόμενο BU.  
**Σκοπός:** να διατηρηθούν οι ακριβείς απαιτήσεις, η κατάστασή τους και οι ανοιχτές επιλογές. **Αυτή η καταγραφή δεν αποτελεί εξουσιοδότηση για νέα native εγκατάσταση.**

## Α. Κλειδωμένες λειτουργικές απαιτήσεις

1. **Απόλυτα συγκρίσιμη κλίμακα μεταξύ θεραπευτών:** στην ίδια περίοδο και στο ίδιο είδος γραφήματος, το κάθε chart έχει το ίδιο y-axis min/max, ακόμη και όταν αλλάζει η επιλεγμένη γραμμή θεραπευτή. Το άνω όριο παράγεται από **τη μέγιστη σχετική τιμή όλων των θεραπευτών της συγκρίσιμης περιόδου**, συμπεριλαμβανομένων και των προγραμματισμένων/πραγματοποιημένων, με ένα σταθερό, ευανάγνωστο άνω περιθώριο εφόσον χρειάζεται. Π.χ. «Συνολικά»: δεν επιτρέπεται Αργέντος με άξονα 0–80 και Καλυβιώτης με 0–100+ όταν προβάλλονται εναλλάξ. Οι επιλεγμένες τιμές και οι δείκτες Μ.Ο. **δεν** πρέπει να αναβαθμονομούν τον άξονα ανεξάρτητα.
2. **Σταθερά γραφήματα κατά το scroll:** η περιοχή πάνω από τη λίστα θεραπευτών μένει ορατή (παγωμένες γραμμές / σωστή διάταξη), κατά το πρότυπο της αναφοράς ασθενών.
3. **Οπτική επισήμανση επιλεγμένου θεραπευτή:** ολόκληρη η επιλεγμένη γραμμή αποκτά έντονο, σαφές χρώμα αντίστοιχο του ενεργού κουμπιού προβολής, η προηγούμενη επιστρέφει στην κανονική εμφάνιση. Η αλλαγή θεραπευτή ενημερώνει τα γραφήματα.
4. **Μ.Ο. ως σημάδι επάνω στην μπάρα, όχι επιπλέον μπάρα:** λεπτή, **ευθεία, συνεχής οριζόντια γραμμή**, στο ύψος του μέσου όρου, **ελάχιστα πλατύτερη** από την αντίστοιχη μπάρα, με μικρή προεξοχή δεξιά/αριστερά. Ο δείκτης δεν πρέπει να καταλαμβάνει όλο το πλάτος του γραφήματος. Ο μέσος όρος προκύπτει από τους συγκρίσιμους θεραπευτές της ίδιας περιόδου. Να παραμένει εμφανής και η αριθμητική τιμή Μ.Ο. χωρίς περιττή πληροφορία.
5. **Χρώμα του δείκτη Μ.Ο.:** η τελευταία διατυπωμένη προτίμηση είναι **κόκκινος κανονικά**, και **φωσφοριζέ πράσινος μόνο αν πέφτει πάνω σε κόκκινη μπάρα**. Ειδικότερα, προσοχή στην περίπτωση θεραπευτή με κόκκινο φόρτο. Να αξιολογηθεί η αντίθεση στην πράξη και στα δύο γραφήματα, όχι μόνο στον κώδικα.
6. **Συνοπτική εικόνα τμήματος:** το «Συνολικά» **ως mode** δείχνει ιστορικό ανά θεραπευτή σε όλη τη διαθέσιμη περίοδο· αυτό διαφέρει από την επιπλέον γραμμή **«ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ»**, η οποία πρέπει να εμφανίζει αθροιστικά τη συγκεκριμένη συγκρίσιμη περίοδο:
   `ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ | Προγραμματισμένες: 434 | Πραγματοποιημένες: 359 | Υλοποίηση: 82,7%`.
   Τύπος: `100 × πραγματοποιημένες / προγραμματισμένες`, 1 δεκαδικό, υπολογισμός από τα **συνολικά πλήθη**, όχι ως μέσος όρος ατομικών ποσοστών. Μηδενικός παρονομαστής → σαφής ένδειξη «μη διαθέσιμο», ποτέ αυθαίρετα 0% ή διαίρεση με μηδέν.
7. **Στόχος χρηστικότητας:** διοικητής/μη τεχνικός χρήστης να καταλαβαίνει την εικόνα αμέσως. Υπολογιστική πολυπλοκότητα στο παρασκήνιο, ελάχιστη αλλά αξιόπιστη πληροφορία στην πρώτη οθόνη.
8. **Ορθός διαχωρισμός στατιστικών:** φόρτος θεραπευτή = προγραμματισμένες, πραγματοποιημένες και πράγματι αχρησιμοποίητες θέσεις με χρονικό/χωρητικό έλεγχο, εφόσον είναι επαληθεύσιμες. Αιτία απώλειας θεραπείας (απουσία ασθενούς, αναβολή ΦΘ / Ρομποτικών / και των δύο, αναστολή) έχει μεγαλύτερη σημασία στη μεριά ασθενούς. **Δεν** συνάγουμε «τελικά κενές θέσεις» απλώς από τις υπάρχουσες μετρήσεις `released` ή `suspensions`, επειδή ενδέχεται να έγινε αναπλήρωση.
9. **Ιστορική ακεραιότητα:** τα πραγματικά συμβάντα απουσίας/αναβολής ασθενούς δεν εξαφανίζονται όταν το κενό καλύπτεται από άλλο ασθενή. Το H9 attendance demo είναι συνθετικό, **όχι** επίσημο πραγματικό ιστορικό ή παραγωγικό KPI. Το τρίτο γράφημα παρουσιών θεραπευτών πρέπει να δηλώνει την έλλειψη τεκμηριωμένης βάσης μέχρι να συνδεθεί με επαληθευμένο ημερολόγιο εργασίμων.

## Β. Νέες οπτικές απαιτήσεις, προς κοινή υλοποίηση και αποδοχή

### Χρώματα μπαρών, **αποκλειστικά δοκιμαστικά**, ΟΧΙ οριστική επιλογή

Ο χρήστης πρότεινε να **δοκιμαστούν** τα ακόλουθα ακριβή HEX:

| HEX | Σημείωση |
|---|---|
| `#FFF200` | έντονο κίτρινο |
| `#5BFA74` | φωτεινό πράσινο |
| `#FFB400` | πορτοκαλοκίτρινο |
| `#FF0303` | φωτεινό κόκκινο |
| `#3FA8F5` | γαλάζιο |

**ΚΡΙΣΙΜΟ:** δεν έχει εγκριθεί οριστικά ποιο χρώμα αντιστοιχεί σε ποιες μετρήσεις/ζώνες ή αν όλη η παλέτα θα υιοθετηθεί. Οποιαδήποτε αντιστοίχιση που προτάθηκε από τον βοηθό (π.χ. 0=μπλε, 1–2=πράσινο, 3–4=κίτρινο, 5=πορτοκαλί, ≥6=κόκκινο) είναι **πρόταση για δοκιμή**, όχι απόφαση του χρήστη. Να αξιολογηθεί μαζί του μετά τη native οπτική δοκιμή, χωρίς να παγιωθεί ως επιχειρησιακή σύμβαση.

### Κάτω από κάθε κύριο γράφημα: μικρή γραμμή τάσης τύπου μετοχής (sparkline)

- Να υπάρχει **ξεχωριστός χαμηλός χώρος κάτω από το αντίστοιχο κύριο γράφημα**, μέσα στα οριζόντια όρια του γραφήματος, όπου εμφανίζεται μικρή χρονοσειρά.
- **Φόρτος:** χρονοσειρά φόρτου με κατακόρυφο εύρος `0–ΧΧ` που προσαρμόζεται στις πραγματικά συγκρίσιμες τιμές και **κοινό y-axis ανά θεραπευτή/ίδια περίοδο**. Να μην ορίζεται αυτόματα από τον επιλεγμένο θεραπευτή.
- **Χρόνος:** από την ημερομηνία έναρξης του σχετικού διαθέσιμου ιστορικού έως σήμερα ή το σχετικό όριο της επιλεγμένης προβολής. Ακριβής κανόνας ανά Ημερήσια/Εβδομαδιαία/Μηνιαία/Συνολικά **χρειάζεται κοινή οπτική απόφαση**, όχι αυθαίρετο κλείδωμα από τον βοηθό. Παλαιότερη πρόταση βοηθού: ημερήσια → έως επιλεγμένη ημέρα, εβδομαδιαία → εντός εβδομάδας, μηνιαία → εντός μήνα, συνολικά → όλο ιστορικό. Αυτό **δεν έχει εγκριθεί**.
- **2ο γράφημα θεραπειών:** να αποσαφηνιστεί μαζί με τον χρήστη αν το mini-trend απεικονίζει πλήθος πραγματοποιημένων, ποσοστό υλοποίησης ή δύο σειρές. Η επιλογή `0–100%` ήταν προηγούμενη πρόταση βοηθού, **όχι τελική απαίτηση**.
- **3ο γράφημα παρουσιών:** να μην επινοηθεί χρονοσειρά από συνθετικές ή μη επαληθευμένες πληροφορίες· διατήρηση unavailable placeholder μέχρι πραγματικό proof.
- Τα sparklines δεν αποτελούν ακόμη ολοκληρωμένη/δοκιμασμένη native δυνατότητα.

## Γ. Στάδιο υλοποίησης: μη συγχέετε source, Excel install και οπτική αποδοχή

- **BO:** επαληθευμένο συνθετικό ιστορικό (25 ημέρες, 522 συνεδρίες, 150 therapist-day, 150 attendance facts, 25 finalized events), immutable βάση για δοκιμές.
- **BL:** προγενέστερη αποδεκτή H8 οθόνη. **BM, BN, BP, BQ:** παγωμένα αποτυχημένα υποψήφια. **BR:** native H10 εγκατάσταση αποδεδειγμένη. **BS:** native install + ανεξάρτητοι protection/H7 UI audits GREEN· οι οπτικές προτιμήσεις χρήστη συνεχίζουν να αλλάζουν.
- **BT:** αλλαγές **μόνο στον κώδικα GitHub**, για κοντό δείκτη Μ.Ο. και ποσοστό υλοποίησης. **Δεν έχουν αναφερθεί source tests ή native δημιουργία BT**. Δεν αποκαλούμε το BT ολοκληρωμένο/αποδεκτό.
- **BU:** συζητήθηκε ως πιθανό όνομα επόμενης δοκιμής για χρώματα + sparklines + κοινές κλίμακες, αλλά **δεν έχει δημιουργηθεί ούτε δεσμευθεί με κώδικα**. Απαιτείται νέο source implementation και tests πριν από οποιαδήποτε native εντολή.
- **Σημαντικό πρόβλημα source που πρέπει να αντιμετωπιστεί:** η τρέχουσα δοκιμαστική έκδοση BT `AddTherapistMeanMarker` τοποθετεί σχήμα σε συντεταγμένες plot area. Δεν έχει επιβεβαιωθεί εγγενώς ότι το Excel το αποδίδει όπως επιδιώκουμε. Επίσης η κοινή κλίμακα μεταξύ θεραπευτών **δεν έχει ακόμη εφαρμοστεί** επαληθευμένα σε όλα τα mode/plots.
- **Ασφαλής πολιτική:** ποτέ overwrite/repair/reseed υπάρχοντος .xlsm. Εάν έλεγχος αποτύχει, παγώνει ο candidate. Πάντα fresh demo από BO, source regression gate, μία native δημιουργία, read-only history/protection/UI audits, τέλος οπτική αποδοχή χρήστη.

## Δ. Ακριβής σειρά στην επόμενη κοινή συνεδρία

1. Ξεκινάμε με **ανάγνωση αυτού του εγγράφου** και `Documentation/CURRENT_HANDOFF.md`, χωρίς να ζητούμε ξανά τις ήδη διατυπωμένες απαιτήσεις.
2. Επιβεβαιώνουμε την τεχνική κατάσταση των εκκρεμών source-only BT αλλαγών, τρέχουμε regression tests και διορθώνουμε κώδικα μόνο αν χρειάζεται.
3. Ενσωματώνουμε **πρώτα κοινή κλίμακα με max από όλη την ομάδα** και επαληθεύουμε ότι αλλαγή θεραπευτή δεν αλλάζει y-axis. Κατόπιν δοκιμαστικά HEX.
4. Σχεδιάζουμε και υλοποιούμε mini-trends σε πραγματικό διαθέσιμο ιστορικό, με ακριβή ορισμό χρονικού παραθύρου και μετρικής ανά κάρτα. Ρωτάμε τον χρήστη **μόνο** αν απαιτείται ουσιαστική απόφαση.
5. Κάθε νέο candidate απαιτεί προηγουμένως source tests GREEN και ελεύθερο νέο όνομα. Ο χρήστης εκτελεί native Excel στο Windows περιβάλλον του. Ελέγχουμε ανεξάρτητα αποθήκευση, προστασία, πλοήγηση και τελική εικόνα.
6. Κλείνουμε τις αλλαγές μόνο όταν ο χρήστης δει ότι η σύγκριση θεραπευτών είναι δίκαιη, τα νέα χρώματα εξυπηρετούν και οι τάσεις παραμένουν αναγνώσιμες.

**ΚΑΤΑΣΤΑΣΗ ΣΥΖΗΤΗΣΗΣ:** ο χρήστης ζήτησε ρητά να σταματήσουμε τώρα την υλοποίηση, **να διατηρήσουμε πλήρεις σημειώσεις και οδηγούς** και να συνεχίσουμε μαζί αργότερα. Δεν εξουσιοδοτείται νέα εγκατάσταση ή δοκιμή σε αυτή τη συνεδρία.

## Νεότερη απόφαση χρήστη — δύο σειρές θεραπειών / αμετάβλητη σημασία χρωμάτων
Ο χρήστης επέλεξε ρητά: **ΜΠΛΕ = ΠΡΟΓΡΑΜΜΑΤΙΣΜΕΝΕΣ, ΚΟΚΚΙΝΟ = ΠΡΑΓΜΑΤΟΠΟΙΗΜΕΝΕΣ**. Η μικρή ιστορική τάση **κάτω από το γράφημα θεραπειών** πρέπει να έχει **δύο γραμμές**, μία από κάθε πλήθος συνεδριών (όχι μία γραμμή ποσοστού, ούτε μία μόνο σειρά πραγματοποίησης). Η ίδια χρωματική σημασία θα ισχύει στις δύο κύριες μπάρες θεραπειών για συνεπή ανάγνωση. Συγκεκριμένα, προτεινόμενες ακριβείς υφιστάμενες χρωματικές τιμές: planned `#3FA8F5`, delivered `#FF0303`. Αυτό είναι χρήστης-επιλεγμένη **σημασιολογική αντιστοίχιση**· παραμένει δοκιμαστική η γενική πεντάχρωμη παλέτα **φόρτου** και τα όριά της.

Source-only commits: `9beeba3` άλλαξε τα 2 βασικά σημεία `series.Points` στο `Python/rehab_excel/therapist_history_report_vba.py`, `a58ea7b` προσαρμόζει regression assertions. **Δεν έχει ακόμα υλοποιηθεί το μικρό ιστορικό δίγραμμο trend**, ούτε έγιναν tests/native Excel σε αυτή τη φάση. Δεν αλλάζει το frozen BO/BS. Προς υλοποίηση: πραγματικές περιοδικές μετρήσεις planned/delivered από verified finalized history, δύο σύντομες lines με ίδια κοινή κλίμακα θεραπευτών· μηδενισμοί/κενά και χρονική οριοθέτηση να χειρίζονται τίμια. Μη δηλώνεται έτοιμο μέχρι να περάσουν tests και native.

## 2026-11-08 — BU source mini trends implemented, awaiting source tests and native Excel proof

Operator directed `μη με περιμενεις, προχωρα`. Continued **without confirmation prompts**. Implemented source and tests on `work/history-statistics-rebuild` only. No preview workbook was created, repaired or overwritten.

**Feature**: two slender historical mini charts beneath main therapist charts, within their A:D / E:H horizontal bounds, sourced **only** from finalized daily therapist history and CURRENT_FINAL patient facts:
- `b558d45`, `2c5c838`: `Python/tools/export_therapist_history_ui.py` preserves existing 15-field report TSV and optionally exports `--trend-output` as UTF-8-SIG 5-column `Date/Therapist/Planned/Delivered/UniquePatients` TSV, ISO-date sorted. Same-source collisions stop before writing, no interpolation or synthetic attendance masquerading as real.
- `df08dd1`, `f771eef`, `fbd0ce9`, `6141339`, `d39810e`: `Python/rehab_excel/therapist_history_report_vba.py` now requests both TSVs in the same exporter call, validates header and records, clears prior hidden Q:U daily trend cache, stores sorted finalized facts, retains current mode in hidden P2, and hides P:U. On selection rebuilds TWO compact charts beneath corresponding bar charts. First: daily distinct patient load; second: **two lines BLUE `#3FA8F5` = programmed sessions, RED `#FF0303` = delivered**. Lines reflect finalized fact dates and chart caption shows earliest/latest available period; fewer than two recorded days → no fabricated line and an insufficient-data label. Trial bars and shared primary axes preserved. Daily mode through selected date, weekly selected week, monthly selected month, overall all finalized history; no selected therapist shows all finalized dates and an aggregated team series; time windows are **provisional design choices** for review.
- Trend y-axis for selected therapists is based on the maximum values among **all therapists within the same time scope**, not selected person's own data; group aggregate view can expand its session axis to fit group totals. Blue/red meaning also applies to the main treatment bars.
- `782d0f8`: short mean markers are red by default; fluorescent green ONLY when the reference line falls on a RED filled bar, including delivered bar and workload bar, not merely because a bar is red.
- `78e725b`, `f46e2fe`, `532486d`, `e4e4845`: update visual, selection/marker and trend contract regressions to distinguish data trends from prohibited full-chart mean reference lines.
- `9c9df4e`: new `Tests/test_h10_finalized_trend_export.py` checks actual sorted daily TSV, uniqueness, unchanged main TSV and collision safety through mocked daily report facts.
- `b85c481`: BU native installer proof contract requires persisted new trend VBA markers; source BO frozen / BU fresh-only gate as before.

**Do not claim source tests or Excel native success.** The GitHub PR-specific workflow status for a recent commit returned no workflows and the assistant environment cannot directly execute Windows Excel. The repo is not locally accessible here for pytest. Need to run on Windows after fast-forward pull: `Tests/test_h10_therapist_simple_demo.py`, `Tests/test_h10_finalized_trend_export.py`, `Tests/test_h8_synthetic_demo_history.py`, `Tests/test_h8_synthetic_history_verifier.py`, `Tests/test_h8_dashboard_visual_navigation.py`, `Tests/test_h8_unique_patient_workload.py`. **If any source fail, STOP; fix source and never create BU.** If green, BO read-only verifier; ensure BU file absent, one native `Python/tools/prepare_h10_therapist_simple_demo.py --source ...BO.xlsm --output ...BU.xlsm`, stop and freeze on any failure, then verify BU read-only history/protection/toolbar, lastly manually inspect monthly/overall/daily/weekly charts, common scale and line color/length alignment. No retroactive overwrite of BS or BO. The 3rd attendance chart stays explicitly unavailable until real history.

**Potential native risk:** Excel COM VBA Chart.Shapes.AddLine and compact mini-chart geometry/axis label behavior have not been proven in Excel. If failure, freeze BU and prepare next fresh candidate; no second attempt. Exact color thresholds for **workload** remain trial only; blue planned/red delivered are explicit user-selected semantics. The mini time-window rules were not specifically user-approved; evaluate in visual trial.

## 2026-11-08 — ΕΠΕΙΓΟΥΣΑ ΣΥΝΕΧΕΙΑ ΜΕΤΑ ΑΠΟ ΕΙΚΟΝΑ BU ΚΑΙ ΝΕΕΣ ΑΠΑΙΤΗΣΕΙΣ ΧΡΗΣΤΗ

**Ο χρήστης ζήτησε ρητά να διασωθεί η συνομιλία ώστε στην επόμενη σελίδα/συνομιλία να συνεχίσουμε αυτομάτως, χωρίς επανάληψη ερωτήσεων ή «οκ». Αυτή η ενότητα έχει προτεραιότητα από παλιότερες σημειώσεις.**

### Επιβεβαιωμένα από native screenshot BU
- Ο χρήστης άνοιξε `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261108_BU.xlsm` και έδειξε screenshot σε mode **«Συνολικά»** με θεραπευτή **Αργέντος**. Δήλωσε: **«όμορφο καθαρό και φωτεινό»**. Screenshot στον προηγούμενο διάλογο (δεν θεωρούμε ότι υπάρχει στο επόμενο εκτός αν ανεβεί).
- Τρία βασικά panel: ασθενείς/ημέρα (αριστερά), θεραπείες (μεσαίο), μη διαθέσιμο ιστορικό παρουσιών (δεξιά). Κάτω από τα δύο πρώτα υπάρχουν πράγματι πολύ χαμηλά mini-trends. Δεύτερο trend μπλε προγραμματισμένες και κόκκινες πραγματοποιημένες.
- Η μπλε συνοπτική γραμμή έγραφε **ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ | Προγραμματισμένες: 551 | Πραγματοποιημένες: 444 | Υλοποίηση: 80,6%**. Αυτό αφορά **συνθετικό demo**, όχι παραγωγική πραγματικότητα. Οι τιμές **434/359** ήταν προηγούμενο παράδειγμα, όχι τρέχον BU mode.
- BU H7 read-only audit operator-proven `H7 INTEGRATED AUDIT OK; workbook not changed`; ο χρήστης δεν έδωσε όλες τις άλλες εξόδους gates, μην εφευρίσκεις.
- Screenshot: τελικό κύριο θεραπειών Μ.Ο. προγραμματισμένων 91,8 / πραγματοποιημένων 74,0, αλλά ο χρήστης παρατήρησε ότι σε κάποιες ημερήσιες κορυφές του **mini trend** το κόκκινο βρίσκεται πάνω από το μπλε. Χρειάζεται πραγματική σημασιολογική διασταύρωση.

### ΑΚΡΙΒΕΙΣ ΝΕΕΣ ΑΠΑΙΤΗΣΕΙΣ ΧΡΗΣΤΗ, εγκεκριμένες για υλοποίηση
1. **Mini trends πρέπει να αποκτήσουν οριζόντιες διαγραμμίσεις (gridlines) με αριθμημένη y κλίμακα**, για να φαίνεται πραγματική μεταβολή.
2. Πιο ψηλά/ευανάγνωστα τα trends. Το αριστερό «Ασθενείς/ημέρα» μπορεί **να μικρύνει αρκετά ή ακόμα και να αφαιρεθεί αν κερδίσουμε χρήσιμο ύψος**. Για αριστερό διατηρούμε 0–6, βήμα 1 όπου καλύπτει τις πραγματικές τιμές. Για δεξί «θεραπείες» θέλει **διαιρέσεις ανά 1 όπου πρακτικά γίνεται** και να είναι σαφές ότι οι αριθμοί αφορούν **μοναδιαίες συνεδρίες, ΟΧΙ αόριστο ×10**. Αν το συνολικό εύρος είναι μεγάλο, η ομαδοποίηση των ticks πρέπει να είναι εμφανώς σημασμένη.
3. **Συγκεντρωτικό γράφημα «ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ» μέσα στην ίδια σελίδα**, **κάτω από τη γραμμή του εκάστοτε τελευταίου θεραπευτή**, δυναμικά αφού ο αριθμός θεραπευτών αλλάζει. Να απεικονίζει τους συγκεντρωτικούς αριθμούς για Ημερήσια/Εβδομαδιαία/Μηνιαία/Συνολικά ανά mode. Προσοχή: το «Συνολικά» mode και το άθροισμα όλων των θεραπευτών σε μία περίοδο είναι διαφορετικές έννοιες.
4. **ΞΕΧΩΡΙΣΤΟ, καθαρό, ΑΣΦΑΛΩΣ ΚΟΙΝΟΠΟΙΗΣΙΜΟ φύλλο διοίκησης** με ΜΟΝΟ τους συγκεντρωτικούς αριθμούς/μπλε γραμμή/γραφήματα, κουμπιά Ημερήσια/Εβδομαδιαία/Μηνιαία/Συνολικά. Να μπορεί να αποσταλεί **ΜΟΝΟ ΑΥΤΟ**, χωρίς στοιχεία ατομικών θεραπευτών, ασθενών, τεχνικών φύλλων, κώδικα/VBA/hidden δεδομένα. **Απλή αποστολή ολόκληρου .xlsm με ένα ορατό «ΔΙΟΙΚΗΣΗ» φύλλο δεν είναι ασφαλής**. Πρέπει να υλοποιηθεί εξαγωγή σε ΞΕΧΩΡΙΣΤΟ αρχείο αποκλειστικά με whitelisted συνοπτικά δεδομένα και οπτικά στοιχεία, π.χ. standalone .xlsx/PDF, κατόπιν ελέγχου διαρροών. Το απλό νέο φύλλο για θέαση ΔΕΝ είναι από μόνο του export privacy proof.
5. **Εξήγηση γιατί κόκκινο > μπλε κάποιες ημέρες:** επαληθεύτηκε στα Python report μοντέλα ότι `planned_sessions` = αρχικός προγραμματισμός, `final_active_sessions` = τελικές πραγματοποιημένες ενεργές συνεδρίες. Το `TherapistDailyReport.session_delta` είναι `final_active_sessions - planned_sessions`. Μπορεί νόμιμα να είναι θετικό λόγω πρόσθετων/εγκεκριμένων μεταβολών, αλλά **δεν έχει γίνει ελεγκτική ανά ημέρα επαλήθευση των συγκεκριμένων σημείων Αργέντος**. Μην «μακιγιάρεις» τα γεγονότα ούτε ισχυριστείς ως αποδεδειγμένο ότι τα συγκεκριμένα σημεία είναι σωστά. Οι ετικέτες/legend θα λένε **«Αρχικά προγραμματισμένες» (μπλε)** και **«Τελικά πραγματοποιημένες» (κόκκινο)** και θα εμφανίζεται εξήγηση όταν actual > planned. Αν χρειάζεται νέα metric «τελικά προγραμματισμένες», πρώτα ορίζουμε/τεκμηριώνουμε από ποια δεδομένα θα παράγεται.
6. **Παλαιότερες απαραίτητες προδιαγραφές εξακολουθούν:** κοινός άξονας y ανά θεραπευτή, ανά ίδιο mode/period από μέγιστο ολόκληρης ομάδας, thin short mean marker κόκκινο/νεον πράσινο πάνω σε κόκκινη μπάρα, blue selected row, sticky charts, πεντάχρωμη παλέτα ΦΟΡΤΟΥ ΜΟΝΟ ΔΟΚΙΜΑΣΤΙΚΗ. Attendance chart unavailable χωρίς πραγματικό validated attendance/calendar.

### Κατάσταση κώδικα αυτή τη στιγμή (χωρίς υπερβολές)
- Στο `work/history-statistics-rebuild` έχει γίνει source-only commit **`fc5ad04119ea09795776c3d36fee29fc2d292511`** στο `Python/rehab_excel/therapist_history_report_vba.py`: mini-trend gridlines, λίγο μεγαλύτερο ύψος τάσης 90 αντί 72, κύρια charts 100 αντί 108 για να χωρέσουν, ορατά axis tick labels, βήμα 1 για αριστερό, βήμα συνεδριών έως περίπου 8 διαιρέσεις για δεξί, Q ISO date column ως text.
- Η επόμενη προγραμματισμένη αλλαγή για συγκεντρωτικό γράφημα **ΔΕΝ ΕΙΧΕ ολοκληρώσει tool call/commit όταν ο χρήστης παρενέβη με αίτημα διαφύλαξης ιστορικού**. **ΜΗΝ** ισχυριστείς ότι υπάρχει ήδη `BuildDepartmentAppendix` ή `MANAGEMENT_SUMMARY` πριν ελεγχθεί πραγματικά στο repo.
- Καμία νέα BV/νέα native δοκιμή από αυτές τις αλλαγές. **Το υπάρχον BU είναι native visual evidence και πρέπει να παραμείνει αμετάβλητο/παγωμένο**. Επόμενος candidate μόνο φρέσκο, με νέο διαφορετικό όνομα **BV (υπόθεση εφόσον δεν υπάρχει)** και green source regression gate πριν από οποιοδήποτε Windows Excel native run.
- Προηγούμενες πηγές: `Documentation/H10_THERAPIST_CHARTS_USER_REQUIREMENTS_20261108.md`, `Documentation/H10_NEXT_SESSION_READ_FIRST_20261108.md`, `Documentation/CURRENT_HANDOFF.md`, `Documentation/PROJECT_STATUS.md`. Αυτά είναι τα authoritative έγγραφα. Συνεχίζουμε εργαζόμενοι αυτόνομα και ενημερώνουμε συνεχώς docs.

### Επόμενη τεχνική σειρά χωρίς άλλη άσκοπη έγκριση
(α) Επανέλεγχος `fc5ad041` GitHub source και αποφυγή native δοκιμής πριν tests· (β) στοχευμένος semantic diagnostic `planned vs actual` ανά ημερομηνία/θεραπευτή με audit for positive deltas; (γ) δυναμικό επιβεβαιωμένα-ορθό department summary chart κάτω από τελευταίο therapist row και κουμπιά mode; (δ) διοικητικό φύλλο μέσα στο workbook ως preview/λειτουργικό UI, ΑΛΛΑ ξεχωριστός **safe sanitized export** μόνο whitelisted δεδομένων σε independent .xlsx/PDF, με tests για absence of names/identifiers/hidden sheets/code and no external links; (ε) regression + independent read-only audit + new unique candidate native proof + κοινή αξιολόγηση εικόνας. **Μην περιμένεις «οκ» από χρήστη**.

## ΕΝΕΡΓΗ ΣΥΝΕΧΕΙΑ H10 / BV μετά το «back to work» και «μη σταματάς» (2026-11-08)

**Υποχρεωτική λειτουργική οδηγία από τον χρήστη:** όταν δεν απαιτείται ουσιαστική επιχειρησιακή απόφαση, **μην περιμένεις άλλο «οκ»**. Προχώρα αυτόνομα σε source, tests, docs και επόμενη ασφαλή πρόταση native proof. Όχι όμως χωρίς τα Windows Excel safety gates.

**BU native baseline:** user screenshot «Συνολικά»/Αργέντος, αποδεκτή φωτεινή απλή αισθητική· **BU παγωμένο**. Η νέα υποψήφια έκδοση είναι αποκλειστικά **φρέσκο BV από frozen BO**. Τα πέντε χρώματα φόρτου παραμένουν ΔΟΚΙΜΑΣΤΙΚΑ. Μπλε αρχικά προγραμματισμένες, κόκκινο τελικά πραγματοποιημένες είναι ρητή επιλογή. Τρίτο panel attendance παραμένει ΜΗ ΔΙΑΘΕΣΙΜΟ.

### Ολοκληρωμένες source-only αλλαγές (ΔΕΝ αποτελούν native proof)
1. `Python/tools/audit_therapist_planned_vs_actual.py` (commit `380e05d`) και `Tests/test_h10_planned_vs_actual_audit.py` (`62ab8e1`): read-only ημερήσιο audit `final_active_sessions > planned_sessions` με ημερομηνία/θεραπευτή/αρχική/τελική/θετική διαφορά/αναπληρώσεις/εγκεκριμένες προσθήκες/αποδεσμεύσεις/αναστολές. Δεν αποδίδει αυτομάτως αιτία· οι συγκεκριμένες κορυφές του Αργέντου απαιτούν BO/BV read-only δεδομένα. **Δεν κόβουμε τεχνητά κόκκινη γραμμή**.
2. `Python/rehab_excel/therapist_history_report_vba.py` commits `fc5ad041` (από προηγούμενη συνομιλία), `8edfd24`, `389f6f8`, `b3ee5f0`, `edde860`, `1a54be5`, `baac9fe`: mini trends πιο ψηλά, ορατές y-gridlines / tick labels, load αριστερά unit 1, session δεξιά unit 1 όταν κλίμακα το επιτρέπει, διαφορετικά ρητή αριθμημένη μεγαλύτερη μονάδα. Ασφαλέστερη σειρά δημιουργίας data series πριν formatting άξονα. Labels **Αρχικά προγραμματισμένες** / **Τελικά πραγματοποιημένες**, ένδειξη πλήθους ημερών με τελικά > αρχικά. Δυναμικό `BuildDepartmentAppendix` κάτω από τελευταίο θεραπευτή, group-by mode period, με planned/delivered totals, ratio από counts, συγκεντρωτικό chart. Οριοθέτηση τελευταίας θεραπευτικής γραμμής σε hidden `P3` για να μην επιλέγεται αναφορά ως άτομο. **Σημαντικό διορθωμένο σφάλμα `baac9fe`**: στην BU η μπλε γραμμή «ΟΛΟΙ» ενίοτε υπολόγιζε μόνο τους θεραπευτές του ίδιου χρόνου με τον επιλεγμένο (π.χ. 551/444). Πλέον το all-therapist visible-mode total είναι **ανεξάρτητο από benchmark cohort/selection**, που χρησιμοποιείται μόνο για Μ.Ο. και κοινό y scale. Το νέο period-level appendix δίνει και λεπτομερή σύνολα.
3. **Εσωτερική διοικητική καρτέλα `MANAGEMENT_SUMMARY`** δημιουργείται και ενημερώνεται από VBA `RenderManagementSummary` με **μόνο συγκεντρωτικά counts** και chart, κουμπιά Ημερήσια/Εβδομαδιαία/Μηνιαία/Συνολικά, μετάβαση σε πλήρη αναφορά, και επιλογή «Εξαγωγή μόνο διοίκησης». **Το εσωτερικό φύλλο ΔΕΝ είναι από μόνο του ασφαλές για αποστολή· ποτέ μην στείλεις το λειτουργικό XLSM**. Το installer `Python/tools/prepare_h10_therapist_simple_demo.py` ενημερώθηκε commit `9432c06` για **BV only** και νέο guard worksheet event, καθώς και `9a4069a` για νέα προστατευμένη management sheet με reopen verification.
4. `Python/tools/export_management_therapist_summary.py` (`369850f`, `2e10cb1`): δημιουργεί **νέο, ανεξάρτητο**, macro-free .xlsx σε Windows Excel από **μόνο αριθμητικά sums** που διαβάζονται από το authoritative current-final report. Τέσσερις καρτέλες mode, εσωτερικά links σαν mode buttons, κάθε καρτέλα δικό της chart και συνολικό ποσοστό. Strict no-overwrite, package inspection ακριβώς τεσσάρων sheet XML, καθόλου vbaProject, external rels/links, hidden sheets, embeddings, queries, schema identifiers. Οι ατομικές εγγραφές/θεραπευτές/ασθενείς δεν μεταφέρονται. **Μόνο μετά native export audit και επιθεώρηση μπορεί να κοινοποιηθεί το νέο XLSX**. Δεν υπάρχει native verified εξαγωγή ακόμη.
5. Tests: `Tests/test_h10_management_export.py` (`abcb56f`, `07f0abd`), `Tests/test_h10_therapist_simple_demo.py` (`567ff7b`, `f534abe`, `a5e1921`, `e95ac09`), `Tests/test_h8_dashboard_visual_navigation.py` (`90f8439`). **Δεν εκτελέστηκε pytest ακόμη**. Έγινε source static inspection μέσω GitHub των απαιτούμενων installer VBA markers, all-therapist totals, colors, 4 management tabs: static checks pass, **δεν υποκαθιστούν pytest**.

### Επόμενη αυστηρά ασφαλής εκτέλεση (ο operator έχει Windows PC)
Με Excel κλειστό, `git pull --ff-only origin work/history-statistics-rebuild`. Τρέχει focused:
`python -m pytest Tests/test_h10_planned_vs_actual_audit.py Tests/test_h10_management_export.py Tests/test_h10_therapist_simple_demo.py Tests/test_h10_finalized_trend_export.py Tests/test_h8_dashboard_visual_navigation.py Tests/test_h8_unique_patient_workload.py -q`.
**STOP** αν όχι GREEN, διορθώνουμε μόνο repo (όχι workbook). Μετά read-only `python Python/tools/audit_therapist_planned_vs_actual.py --workbook Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm --therapist Αργέντος`; no data mutation. Αφού ελεγχθούν και το BO snapshot verifier και ελεύθερο path, μία μόνο native δημιουργία `BV`:
`python Python/tools/prepare_h10_therapist_simple_demo.py --source Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm --output Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261108_BV.xlsm`.
Κατόπιν independent read-only history/protection/H7 audits, native visual/chart/management navigation tests. Για share test χρησιμοποιούμε **άλλο νέο** .xlsx path και `python Python/tools/export_management_therapist_summary.py --source ...BV.xlsm --output ...MANAGEMENT_BV_DEMO.xlsx`. Οποιοδήποτε failure ⇒ freeze output και stop, όχι second attempt.

**Προσοχή σε δύο άλυτα risks:** macro procedures/Excel ChartObjects/axes δεν δοκιμάστηκαν ακόμη στη BV και οι προηγούμενες μέσες τιμές στο screenshot χρειάζονται retest· ορισμένα αρχεία `CURRENT_HANDOFF.md` έχουν παλιές ιστορικές σημειώσεις που **δεν** αντικαθιστούν αυτή τη νεότερη καταγραφή. Ο χρήστης θέλει συνεχή documentation για seamless μεταφορά σε νέα συζήτηση.
