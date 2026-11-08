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
