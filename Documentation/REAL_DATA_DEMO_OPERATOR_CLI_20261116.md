# DIOMIL: πραγματικά δεδομένα και τοπική προετοιμασία DEMO (2026-11-16)

**Τρέχουσα κατάσταση:** Υπάρχει πλέον **εκτελέσιμο CLI εργαλείο προελέγχου**, όχι ακόμη έτοιμο πραγματικό XLSM. Η δημιουργία με Microsoft Excel/COM δεν έχει εκτελεστεί ή επικυρωθεί στο ιδιωτικό Windows σύστημα. **Το demo αποδοχής απαιτεί πραγματικούς ασθενείς/θεραπείες/ώρες/θεραπευτές**, όχι συνθετικούς.

## Κώδικας

`Python/tools/prepare_real_data_practice_demo.py`

Διαθέτει δύο διακριτές λειτουργίες:
1. **Default READ-ONLY**, χωρίς `--build`: διαβάζει το πραγματικό χειροκίνητο A:V, το εγκεκριμένο patient/weekly review, το κενό επιλέξιμο XLSM και το *ανεξάρτητα εγκεκριμένο* ιδιωτικό JSON κλινικών στοιχείων. Επαληθεύει source hashes, πλήρη weekday/time/therapist reconciliation, identity review, F/U/V blockers, κενό πραγματικό VBA shell και συνοχή μητρώου. **Δεν γράφει XLSM**.
2. **Ρητό `--build`** (Windows + εγκατεστημένο Excel/pywin32): ξανατρέχει όλες τις ίδιες πύλες και κάνει native seed σε *νέο, ανύπαρκτο* `*_DEMO_*.xlsm`, **εκτός Git**, με ανεξάρτητο readback. Δεν αντικαθιστά ή γράφει την πραγματική χειροκίνητη πηγή. Το XLSM που παράγεται είναι μόνο υποψήφιο προς πρακτικό έλεγχο.

Τα πεδία/όνομα/ID/πηγαία cells/ιατρός/MRN δεν τυπώνονται. Αποτυχία εκτυπώνει μόνο **κλάση σφάλματος**. Τα προσωπικά δεδομένα δεν αποθηκεύονται σε GitHub, CI logs ή κοινόχρηστους φακέλους.

## Απαιτούμενα ιδιωτικά αρχεία

- `manual_current.xlsx`: **φρέσκο snapshot** της τρέχουσας χειροκίνητης πηγής για την εβδομάδα σύγκρισης. Όχι αυθαίρετα η παλιότερη έκδοση 10/11/2026.
- `review_confirmed.xlsx`: ιδιωτική επαληθευμένη πηγαία χαρτογράφηση/συσχέτιση IDs και εγκεκριμένα ακριβή εβδομαδιαία source placements, υπό το schema του `trial_admission_gate.py`.
- `eligible_empty_shell.xlsm`: πλήρες, αλλά **εντελώς κενό** κανονικό Diomil Excel με VBA, όλα τα κουμπιά/φύλλα και κανένα προηγούμενο ιστορικό/ασθενή.
- `clinical_approvals.json`: **ιδιωτική, πλήρης, ρητά εγκεκριμένη** κλινική ταξινόμηση ανά (`clinic`, `source_row`) του ίδιου review. Μην αντιγράφετε μόνο ονοματεπώνυμα και υποθέτετε κλινικές πληροφορίες.

Το JSON έχει ακριβώς την εξής σύνταξη, με **αποκλειστικά φανταστικές τιμές στο παράδειγμα**. Οι δύο ενδεικτικοί ασθενείς είναι εσωτερικοί, επειδή το τρέχον αρχικό A:V importer υποστηρίζει μόνο εγκεκριμένες εσωτερικές εγγραφές. Τα πραγματικά εξωτερικά προγράμματα απαιτούν ανεξάρτητη πηγή/εισαγωγή πριν θεωρηθεί πλήρης η σύγκριση:

```json
{
  "schema": "diomil-private-clinical-approvals-v1",
  "records": [
    {
      "clinic": "DEMO CLINIC A",
      "source_row": 2,
      "patient_type": "INPATIENT",
      "infectious": false,
      "room": "DEMO A01",
      "responsible_doctor": null,
      "hospital_mrn": null,
      "status": "Ενεργός"
    },
    {
      "clinic": "DEMO CLINIC B",
      "source_row": 3,
      "patient_type": "INPATIENT",
      "infectious": false,
      "room": "DEMO B02",
      "responsible_doctor": null,
      "hospital_mrn": null,
      "status": "Ενεργός"
    }
  ]
}
```

**Το δείγμα JSON δεν είναι έγκριση πραγματικών ασθενών.** Οι πραγματικές `clinic` πρέπει να ταιριάζουν **ακριβώς** με το review, `source_row` είναι ο αριθμός αρχικής γραμμής στο χειροκίνητο πρόγραμμα (πάντα ≥2), `patient_type` μόνο `INPATIENT` ή `OUTPATIENT`, `infectious` **πραγματικό JSON boolean** `true`/`false`, χωρίς αυτόματη συναγωγή από χρώμα. Τα προαιρετικά κλινικά πεδία επιτρέπουν `null`, όχι αυθαίρετες πληροφορίες. Όλοι οι επιβεβαιωμένοι ασθενείς του review πρέπει να καλύπτονται μία φορά, χωρίς διπλότυπα. Οι τιμές παραμένουν τοπικά και ιδιωτικά.

## Εντολές PowerShell

Σε **Windows**, με ενεργό το Python περιβάλλον του repo και εγκατεστημένες εξαρτήσεις, θέστε διαδρομές που δείχνουν σε **ιδιωτικό εξουσιοδοτημένο φάκελο, εκτός Git**. Οι τιμές είναι *placeholders*, ποτέ πραγματικά ονόματα/νοσηλεία σε δημοσιευμένο script.

```powershell
$private = "C:\PRIVATE_APPROVED_DIOMIL"
$programme = Join-Path $private "manual_current.xlsx"
$review = Join-Path $private "review_confirmed.xlsx"
$shell = Join-Path $private "eligible_empty_shell.xlsm"
$approvals = Join-Path $private "clinical_approvals.json"

# Πρώτα: πραγματικός πλήρης προέλεγχος ΧΩΡΙΣ αλλαγή XLSM.
python Python\tools\prepare_real_data_practice_demo.py `
  --programme "$programme" `
  --review "$review" `
  --shell "$shell" `
  --clinical-approvals "$approvals"

if ($LASTEXITCODE -ne 0) {
    throw "STOP: real-data preflight failed; do not build."
}

# ΜΟΝΟ μετά από την απαραίτητη ιδιωτική έγκριση και επιβεβαίωση
# όλων των ανοιχτών F/U/V και προγραμματιστικών conflicts.
$newCandidate = Join-Path $private "DIOMIL_DEMO_PRIVATE_FIRST_A.xlsm"
if (Test-Path -LiteralPath $newCandidate) {
    throw "STOP: this DEMO destination already exists"
}
python Python\tools\prepare_real_data_practice_demo.py `
  --programme "$programme" `
  --review "$review" `
  --shell "$shell" `
  --clinical-approvals "$approvals" `
  --build `
  --output "$newCandidate"

if ($LASTEXITCODE -ne 0) {
    throw "STOP: private native DEMO not verified. Do not reuse output."
}
```

**Οι παραπάνω εντολές δεν έχουν εκτελεστεί στο πραγματικό Windows σύστημα.** Εάν λείπει οποιοδήποτε εγκεκριμένο αρχείο, σταματά η εργασία. Το standard `preflight_trial_excel_sources.py` CLI μπορεί να απορρίψει κενό shell επειδή δεν έχει ασθενείς· εδώ χρησιμοποιείται η στενά ελεγχόμενη `empty_demo_shell=True` εξαίρεση μαζί με τον αυστηρότερο `assert_clean_native_demo_shell`.

## Υποχρεωτική χειροκίνητη αποδοχή μετά το build

Η παραγωγή private XLSM **δεν** σημαίνει ότι λειτουργεί το ημερήσιο σύστημα. Άνοιγμα native Excel, αποδοχή macros μόνον εφόσον έχει εγκριθεί από τον φορέα, ημερήσιο πρόγραμμα/MASTER και αναλυτικό εβδομαδιαίο, **πλήρης 1:1 σύγκριση** με το *ίδιο* χειροκίνητο snapshot ΔΕ–ΠΑ, επιβεβαίωση όλων των ασθενών, θεραπευτών, ειδικοτήτων και ωρών χωρίς απώλειες, απουσίες/αναπληρώσεις/επιστροφές/εισαγωγή/εξιτήριο/φοιτητές/θεραπευτές και Save/Close/Reopen. Η ημερομηνιακή προβολή MASTER είναι επί του παρόντος opt-in και δεν ενημερώνεται πλήρως αυτόματα από το κουμπί εξιτηρίου.

**Ανοιχτά:** F/U/V (30λεπτο ανακλινόμενο, απογευματινό, πισίνα), πηγαία outpatient υποδοχή από την κατάλληλη πηγή, πλήρες real-source agreement, ιδιωτικό approved review/clinical facts, χειροκίνητη native δοκιμή και full end-to-end acceptance. **Μην παρακάμπτετε blockers ούτε κατασκευάζετε ψεύτικες καταχωρίσεις**.

Απόδειξη και συνολικό status: `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`. Αν κάποιο native build αποτύχει, **παγώστε** το αντίγραφο και τα ιδιωτικά διαγνωστικά, μην επαναδοκιμάσετε στο ίδιο output. Ενημερώστε τα live `CURRENT_HANDOFF`, `PROJECT_STATUS`, `PROJECT_DELIVERY_MASTER_REPORT` και acceptance matrix χωρίς τροποποίηση frozen H2–H10 αποδείξεων.


### Έλεγχος αμεταβλητότητας κλινικών εγκρίσεων

Η εντολή ελέγχει και SHA-256 του ιδιωτικού `clinical_approvals.json` πριν/μετά τον πλήρη προέλεγχο και μετά το native build. Αν αλλάξει έστω και ένα byte, κάνει **SAFETY STOP**. Εάν η αλλαγή διαπιστωθεί μετά από ήδη ολοκληρωμένη native δημιουργία, το απομονωμένο αποτέλεσμα **παγώνει και ΔΕΝ γίνεται αποδεκτό**. Δεν επαναχρησιμοποιούμε το ίδιο filename ούτε δημοσιεύουμε τις εγκρίσεις.
