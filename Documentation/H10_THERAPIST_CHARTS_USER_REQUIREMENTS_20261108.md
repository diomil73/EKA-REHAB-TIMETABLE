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
