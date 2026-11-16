# DIOMIL: πραγματικά δεδομένα και τοπική προετοιμασία DEMO (2026-11-16)

## UPDATE 2026-11-16 — approved separate outpatient source required for native build

This section **supersedes older full-build examples below that omit outpatient input**. Existing daily MASTER treatment renderer and approved UI stay unchanged. Source code CI at `70e1b88f6e3289bbc79fe411d4724e327622ff7f` is green (run 43809724956: 2,029 tests each Windows/Linux). No actual private source/native Excel candidate has been executed.

Use only authorized, independently approved PRIVATE files in an on-premises Windows folder outside Git. First run the two separate READ-ONLY safety audits; a clean return from either one is diagnostic only, not proof of clinical admission:

```powershell
$private = "C:\\PRIVATE_APPROVED_DIOMIL"  # illustrative private path, NOT Git
$manual = Join-Path $private "manual_current.xlsx"
$review = Join-Path $private "review_confirmed.xlsx"
$clinical = Join-Path $private "clinical_approvals.json"
$shell = Join-Path $private "eligible_empty_shell.xlsm"
$outpatients = Join-Path $private "outpatients_current.xlsx"
python Python\\tools\\audit_real_manual_specials.py --programme "$manual"
python Python\\tools\\audit_real_outpatient_source.py --source "$outpatients"
```

If special F/U/V (reclining, historical afternoon or pool) are present, **STOP** until separately authorized rules/record mapping and the complete real source are accepted. Never suppress the gate by deleting source cells. When no holds remain and the independent source/identity/clinical approvals are complete, run:

```powershell
python Python\\tools\\prepare_real_data_practice_demo.py `
  --programme "$manual" --review "$review" --shell "$shell" `
  --clinical-approvals "$clinical" --outpatients "$outpatients"
if ($LASTEXITCODE -ne 0) { throw "STOP: private real-source preflight failed" }

$newCandidate = Join-Path $private "DIOMIL_DEMO_PRIVATE_A.xlsm"
if (Test-Path -LiteralPath $newCandidate) { throw "STOP: choose a NEW output" }
python Python\\tools\\prepare_real_data_practice_demo.py `
  --programme "$manual" --review "$review" --shell "$shell" `
  --clinical-approvals "$clinical" --outpatients "$outpatients" `
  --build --output "$newCandidate"
if ($LASTEXITCODE -ne 0) { throw "STOP: DO NOT REUSE the candidate" }
```

**Build success is NOT operator acceptance**. Native Excel COM and 1:1 private five-day admission plus full workbook interaction and Save/Close/Reopen are separately required. Never paste patient IDs, room/doctor details or error dumps into public GitHub, chat, logs or CI. For CLI any exception emits only a generic error class. Operator must use a fresh current manual, not an old snapshot.


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


## 2026-11-16 — F/U/V read-only gate, 1,991 tests GREEN

**Concrete new gate, not a synthetic replacement for the user's real-data DEMO:** `Python/rehab_excel/trial_special_fields_audit.py`, `Python/tools/audit_real_manual_specials.py`, `Tests/test_trial_special_fields_audit.py` provide the first **single-source-only** private read-only inspection of actual manual A:V columns **F** (reclining, known 30-minute ΦΘ-related operation), **U** (legacy afternoon, no auto-import) and **V** (morning pool). No review/clinical approval/clean XLSM required for this *diagnostic* first command. It checks original source sheet/header, SHA-256 before/after, cached and raw Excel formula values (never omit formula-only cells), missing-patient source rows, time parseability, all counts and no PII emitted. The tool **never creates sessions, guesses F/V weekdays or imports old U**. It returns 0 if F/U/V truly empty (not whole source approval), 3 for pending real special records, 2 on safety stop. Exact local command, interpretations and privacy gates: `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md`.

**Verified code evidence:** `95157af7dbb32c6cb8aca210c503ffbb7f63be17`, GitHub Actions `43804737124`, **1,991 passing pytest in both Windows and Linux**, all four full/synthetic CI jobs successful. This is not proof that the real private manual has been scanned or a working native Excel was built. **Real-data DEMO remains NOT DELIVERED.** F/V day/provider/duration-aware scheduling approval, U source scope and independently reviewed outpatient programme remain blockers. Next meaningful step is run the read-only audit against the operator-approved private CURRENT manual snapshot, capture only totals, finish explicit F/V mapping and standalone outpatient source, then eligible private native seed and full side-by-side weekday/operator acceptance. Frozen H2–H10 proof artifacts remain untouched.


## 2026-11-16 — Mandatory F/U/V source gate before real DEMO seed

`Python/tools/prepare_real_data_practice_demo.py` now calls `inspect_real_manual_special_fields(programme)` **before reading approved clinical values, building source plans or invoking native Excel**. Any populated F reclining, U legacy afternoon or V morning pool field (including Excel formula-only content with empty cache), or an orphaned source row, gives a privacy-safe **SAFETY STOP**. This ensures that even a stale review workbook cannot silently hide real original-source special therapies. F/V clock parseability is diagnostic only: **days/duration-aware scheduling and authoritative source treatment review are still required**. U remains expressly non-importable from the historical column; separately entered authorized afternoon appointments belong to their own programme. The standalone command `python Python\tools\audit_real_manual_specials.py --programme "<PRIVATE_CURRENT_MANUAL>.xlsx"` reports only aggregate counts and returns 3 for blocked real specials, 0 if no F/U/V cells are populated, 2 on structural safety failure. No original workbook or native XLSM was changed in the GitHub development environment.

**Tests:** `Tests/test_prepare_real_data_practice_demo.py` adds explicit refusal of preflight and native build before clinical approval/Excel on pending F/U/V/orphan formula cells, alongside `Tests/test_trial_special_fields_audit.py`. Verified integrated hard-stop code checkpoint `a9d2f0afd8b119b69a562475525d30d35d7cd8a3`: **1,993 tests passed on Windows and Linux, 4/4 CI jobs successful** in GitHub Actions run `43805000757`. This is source/synthetic regression evidence, not a real-manual scan or native Excel acceptance. Real private source inspection, source review, native Excel candidate and user acceptance remain **PENDING**. The new hard-stop is deliberate and must NOT be bypassed with invented patients or source omissions.


## 2026-11-16 — Real outpatient source audit reuses existing application schema

**Current verified architecture:** Normal patient registration/program editor already writes `PATIENT_PLANNER` and uses `PatientProgramMasterRender.RefreshMasterPatientProgram` for the existing MASTER cell content. The new bulk real-data DEMO seed is a *different* initial-population path; it does not imply the old MASTER renderer was broken. Do NOT rebuild established cell formatting, controls or clinical treatment categories.

**Next separate real-data source bridge:** `Python/rehab_excel/trial_private_outpatient_audit.py` and `Python/tools/audit_real_outpatient_source.py` inspect only a PRIVATE, institution-approved current outpatient export in `.xlsx` form. They require the already-supported `PATIENTS` and `OUTPATIENT_SCHEDULE` sheets and their existing exact leading column headers. Reuse `read_patient_registry`, `read_outpatient_schedule` and `parse_day_pattern` to check outpatient identity/type/name consistency, Mon-Fri days, times, provider for ΦΘ, no robotic outpatient, per-weekday duplicate bookings, no formula-only or orphan source rows and SHA-256 read-only preservation. CLI emits only aggregate counts; never names/IDs/paths. `Tests/test_trial_private_outpatient_audit.py` uses only fictional private fixtures. **This is a diagnostic admission precondition, not an outpatient import or proof of completeness**. Full native outbound bridging from this separate approved source into the same private XLSM remains PENDING. Do not infer outpatients from the existing inpatient A:V review or omit them from the final pilot. Full command and schema: `Documentation/REAL_OUTPATIENT_SOURCE_EXISTING_SCHEMA_AUDIT_20261116.md`.

**Real workbook and operator acceptance status remains unchanged:** no current actual source files available in GitHub, no private native Excel produced or verified, outpatient source admission/merged seed and F/U/V approvals still open, 1:1 real manual comparison pending. Previous 2,002 test green code checkpoint `3cfdf61b` with GitHub run `43805735643` establishes the earlier MASTER bulk-display integration, not outpatient native acceptance. Check the new outpatient audit CI separately. Frozen H2–H10 acceptance files stay immutable.
