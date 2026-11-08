# READ FIRST — επόμενη κοινή συνεδρία για στατιστικά θεραπευτών

**Ο χρήστης ζήτησε ρητά να καταγραφούν οι απαιτήσεις, ΟΧΙ να υλοποιηθούν σήμερα.** Μην επαναλαμβάνεις ερωτήσεις που έχουν απαντηθεί. Μην ξεκινάς Excel installation χωρίς source GREEN.

## Μοναδικός πλήρης οδηγός προδιαγραφών

**[H10_THERAPIST_CHARTS_USER_REQUIREMENTS_20261108.md](H10_THERAPIST_CHARTS_USER_REQUIREMENTS_20261108.md)**

Αυτό είναι το authoritative αρχείο όσων είπε ο χρήστης, των ανοιχτών αποφάσεων, του διαχωρισμού εγκεκριμένου/δοκιμαστικού και του ασφαλούς πλάνου.

## Γρήγορο checklist

- [ ] **ΚΟΙΝΗ ΚΛΙΜΑΚΑ** ανά ίδια προβολή/περίοδο για κάθε θεραπευτή. Ακόμη και αν αλλάζεις όνομα, ο y-axis παραμένει σταθερός και προκύπτει από global max όλων των συγκρίσιμων θεραπευτών. Παράδειγμα Αργέντος 80 / Καλυβιώτης 100+ απαγορεύεται.
- [ ] **ΔΟΚΙΜΑΣΤΙΚΗ, όχι τελική παλέτα:** `#FFF200`, `#5BFA74`, `#FFB400`, `#FF0303`, `#3FA8F5`. ΜΗΝ παγιοποιήσεις αυθαίρετη αντιστοίχιση ζωνών.
- [ ] **Μ.Ο.** κοντή, λεπτή, συνεχής οριζόντια γραμμή που εξέχει ελάχιστα από τη μπάρα, κόκκινη κατά κανόνα, fluorescent πράσινη πάνω σε κόκκινη μπάρα. Δεν θέλουμε ούτε επιπλέον μπάρα ούτε μακριά μωβ διακεκομμένη γραμμή.
- [ ] **Συνοπτικό σύνολο τμήματος:** προγραμματισμένες, πραγματοποιημένες, **υλοποίηση %** με ένα δεκαδικό από πραγματικά αθροίσματα (π.χ. 434/359 → 82,7%). Δεν αθροίζουμε ποσοστά.
- [ ] **Mini trend κάτω από κάθε βασικό γράφημα**, χαμηλή ξεχωριστή ζώνη εντός πλάτους γραφήματος, σαν stock sparkline. Χρονικό παράθυρο και κατακόρυφη κλίμακα ανά mode, αλλά σταθερή σύγκριση μεταξύ θεραπευτών. Ακριβείς ορισμοί ανά προβολή **χρειάζονται επιβεβαίωση** αφού παρουσιαστούν πρακτικές επιλογές.
- [ ] **Sticky γραφήματα + blue active therapist row**, χωρίς να χάνεται η τρέχουσα επιλογή.
- [ ] **Μη διαθέσιμο 3ο attendance γράφημα** έως τεκμηριωμένο πραγματικό attendance/working calendar. Synthetic data = test/demo only.
- [ ] **Διοικητική απλότητα:** στοιχεία που κατανοούνται αμέσως. Οι αναλυτικές αιτίες απώλειας θεραπείας ανήκουν κυρίως στην αναφορά ασθενών, όχι στη βασική σύνοψη φόρτου θεραπευτή.
- [ ] **Κλείδωμα ασφάλειας:** BO trusted synthetic source, BT source-only/untested, BU merely suggested name, BS prior candidate. Δεν αγγίζουμε frozen/accepted demo. Πάντα νέο workbook και stop σε κάθε αποτυχία.

## Κατάσταση στο κλείσιμο

**Μόνο documentation update.** Όχι νέος κώδικας / νέα native εγκατάσταση από αυτό το checkpoint. Ξεκινούμε την επόμενη συνεδρία εδώ και συνεχίζουμε μαζί τις εκκρεμείς οπτικές/στατιστικές επιλογές.

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
