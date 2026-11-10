# Diomil: ασφαλής υποδοχή επικαιροποιημένου αντιγράφου για παράλληλη πιλοτική λειτουργία

## Απόφαση και προτεραιότητα
Ο χρήστης επέλεξε ορθότητα, ασφάλεια και καθημερινή χρηστικότητα πριν από τελικό αισθητικό polish. Το υφιστάμενο αρχείο τμήματος παραμένει **μοναδική επιχειρησιακή πηγή** έως την ξεχωριστή, ρητή αποδοχή του πιλοτικού. Το Diomil αρχίζει ως ξεχωριστό αντίγραφο για παράλληλη αξιολόγηση χωρίς αυτόματη αμφίδρομη σύνδεση, write-back, συγχρονισμό ή αλλαγές στο τρέχον αρχείο.

Όλα τα παλαιότερα αρχεία Diomil AZ/BO/BY/CA και καταγραφές τους είναι **αρχεία προηγούμενων tests**, όχι επιβεβαιωμένα πραγματικά θεραπευτικά συμβάντα. Τα παγώνουμε για regression/diagnostic use, όχι για παραγωγική μεταφορά ιστορικού. Αποφεύγουμε μεταφορά ασθενών/πραγματικών προσωπικών δεδομένων σε δημόσιο GitHub ή σε debug logs.

## Στάδια εισαγωγής
1. Παραλαβή **αντιγράφου** του τρέχοντος λειτουργικού αρχείου, ποτέ του μοναδικού κύριου. Αρχικό hash και ακριβές όνομα υποψηφίου σε τοπικό, περιορισμένης πρόσβασης χώρο. Εφόσον περιέχει αναγνωριστικά ασθενών, προτιμάται ανωνυμοποιημένο τεχνικό δείγμα για ανάπτυξη.
2. **Read-only intake** `Python/tools/audit_pilot_source_readiness.py --workbook "<copy.xlsm>"`. Εξετάζει μόνο τα βασικά φύλλα, την ύπαρξη VBA, ελάχιστη ακεραιότητα PATIENTS, υφιστάμενα history data και hash unchanged. Εκτυπώνει μόνο πλήθη/issue codes, όχι ονόματα, κωδικούς ασθενών ή ολόκληρες συνεδρίες. Exit code `3` σημαίνει **REVIEW REQUIRED**, όχι εντολή αυτόματης επιδιόρθωσης.
3. Καθορισμός source-to-target mapping και ημερομηνίας έναρξης πιλοτικού. Διαχωρισμός pre-pilot μη επιβεβαιωμένων history/test rows από νέες πιλοτικές καταχωρίσεις με ρητή προέλευση. **Απαγορεύεται** αυτόματο copy/migrate ιστορικού μόνο και μόνο επειδή οι στήλες ταιριάζουν.
4. Δημιουργία **μοναδικού νέου** .xlsm από ελεγμένη πηγή, αποκλειστικά σε ασφαλές περιβάλλον Windows Excel, μόνο αφού εγκατασταθούν/επαληθευτούν τα macro contracts. No overwrite, no retry of failed candidate. Independent read-only schema/hash and protected-session audits.
5. Δοκιμή με ελάχιστα σενάρια: ανοιχτή/κλειστή ημέρα, απουσία ασθενούς, απουσία θεραπευτή, αντικατάσταση, διπλοκράτηση, ακύρωση, rollback, save-and-reopen, αναφορά ίδιων totals. Μόνο έπειτα προχωράμε σε πραγματικό shadow run.
6. Κατά τη διάρκεια shadow run, το νέο αρχείο **δεν καθοδηγεί** πραγματικές κλινικές αποφάσεις. Συγκρίνουμε καθημερινά εξόδους με το κύριο αρχείο· αποκλίσεις καταγράφονται πριν οποιοδήποτε release. Δεν υπάρχει αυτόματο cutover.

## Διαχωρισμός blocking vs deferred
**Blocking πριν shadow pilot:** ασυμφωνία ταυτοτήτων, χαμένο/διπλό session, λάθος slot/capacity, παραβίαση snapshot/reopen/rollback, ασυμφωνία source-to-target, overwrite risk, exposure προσωπικών δεδομένων, λάθος αρχικοποίηση ιστορικού.

**Deferred, εφόσον δεν κρύβουν λάθος δεδομένων:** χρωματική παλέτα, λεπτομέρειες γραφικών, οπτικό polish διοίκησης. Ωστόσο επικαλύψεις που κρύβουν κλινικά/λειτουργικά νούμερα είναι usability BLOCKERS, όχι απλό styling.

## Τρέχουσα κατάσταση (source-only)
- Δημιουργήθηκε `Python/tools/audit_pilot_source_readiness.py` και regression tests `Tests/test_pilot_source_readiness.py`. **Αναμένονται Windows pytest**, δεν έγινε native install.
- Ο χρήστης δεν έχει ακόμη ανεβάσει το επικαιροποιημένο αντίγραφο. Δεν είναι δυνατόν να επικυρώσουμε σήμερα compatibility των πραγματικών στηλών, φόρμουλων και macro names.
- BZ παραμένει unbuilt. Το CA παραμένει frozen test candidate που αποτυγχάνει στα strict data gates λόγω παλαιότερων test history facts.


## 2026-11-09: επιβεβαίωση πηγής και ενίσχυση του intake
Ο χρήστης επιβεβαίωσε **44 passed in 2.23s**, `PILOT READINESS SOURCE TESTS GREEN`. Πρόκειται αποκλειστικά για tests κώδικα απογραφής, όχι για πραγματικό λειτουργικό αντίγραφο.

Ενίσχυση source-only:
- `7c88915`: ανιχνεύουμε historical rows ακόμη και με κενό SnapshotId/στήλη A (παλιότερα παραλείπονταν), ελέγχουμε ακριβείς επικεφαλίδες ιστορικών πινάκων πριν από τυχόν mapping και εντοπίζουμε `xl/externalLinks` / `xl/connections.xml` ως δυνητικές εξαρτήσεις. Κάθε τέτοιο εύρημα είναι `REVIEW REQUIRED` και **όχι** αυτόματη μεταφορά ή διαγραφή.
- `e37caf5`: τρία νέα regression tests για κενό SnapshotId, λανθασμένες επικεφαλίδες και ZIP externalLink part.
- Τα νέα tests **δεν έχουν εκτελεστεί ακόμη στα Windows**. Δεν χρησιμοποιήθηκε/τροποποιήθηκε κανένα Excel, δεν υπάρχει λειτουργικό αντίγραφο από τον χρήστη ακόμη.
- `audit_pilot_source_readiness.py` δεν μπορεί να επιβεβαιώσει VBA runtime, προστασίες, αξιόπιστο save/reopen, χρήση όλων των θεραπευτών, ασφάλεια παραλλαγών, διαθεσιμότητα ωρών ή συμβατότητα καθημερινής λειτουργίας. Αυτά παραμένουν blocking gates με native proof όταν φτάσει το ασφαλές αντίγραφο.


## 2026-11-09: λίστα λειτουργικής αποδοχής με υποχρεωτικά evidence refs
**Ο χειριστής επιβεβαίωσε** `47 passed in 3.46s` / `PILOT INTAKE SAFETY GATES GREEN`. Οι source-only βελτιώσεις inventory/links/headers πέρασαν.

Στο επόμενο checkpoint προστέθηκαν:
- `Python/tools/verify_pilot_acceptance_evidence.py`: schema-versioned JSON checklist με **13 υποχρεωτικές λειτουργικές δικλείδες**: μητρώο και δομή, ιστορική προέλευση, hash και απομονωμένο αντίγραφο, χωρητικότητα, ταυτότητες/πλήθη συνεδριών, απουσίες και ακυρώσεις, απουσία και αναπλήρωση θεραπευτή, διπλοκράτηση/έγκριση, οριστικοποίηση+reopen, amendment/restore/rollback, VBA/protection/navigation, στατιστικά και ημερολογιακή οριοθέτηση, απομόνωση write-back προς το κύριο αρχείο.
- Δημιουργία μόνο με fresh `--create-template <new-local.json>` (exclusive create, ποτέ overwrite), όλα τα tests αρχικά `pending`. Επικύρωση μόνο με `--manifest <local.json>`: 0 όταν η **δήλωση τεκμηρίων** είναι πλήρης, 3 αν υπάρχουν pending/failed/blocked ή λείπουν hashes/native checking και 2 αν είναι άκυρη η δομή. **Ακόμη και 0 δεν συνιστά ανεξάρτητη αποδοχή ή άδεια παραγωγής**, απαιτεί πραγματικό native Excel proof και ανθρώπινο έλεγχο αποτελεσμάτων.
- Το manifest περιέχει **μόνο hashes, ημερομηνία, κωδικούς κατάστασης και σύντομα μη προσωπικά evidence identifiers**. Όχι στοιχεία ασθενών, snapshots/session IDs, ούτε raw Excel μέσα σε GitHub. Δεν επιβάλλεται δημιουργία manifest μέχρι να έρθει το επικαιροποιημένο αντίγραφο.
- `Tests/test_pilot_acceptance_evidence.py` δοκιμάζει pending, valid/invalid contract, no-overwrite, missing proof και ίδια hashes candidate/reference. **Source-only, αναμένει pytest στα Windows**, δεν έγινε native Excel install.


### 2026-11-10 — Real manually maintained workbook, source staging preview (NO IMPORT)
User supplied `ΠΡΟΓΡΑΜΜΑΤΑ  ΘΕΡΑΠΕΙΩΝ ΑΣΘΕΝΩΝ 10-11-26.xlsx`; archived source hash SHA-256 `765264ec8eae95a31ccc57629dc31b1d61f59ea5c992221affa91fcf6a8d2567`. Analysis read the **visible `ΠΡΟΓΡΑΜΜΑ`** worksheet resolved by actual workbook XML relationship (not sheet1.xml); remaining 66 hidden sheets were not construed as current appointments. No changes to uploaded source; hash unchanged after output.
- Real manual interpretation contract from operator: FIRST clock time and FIRST therapist = original planned appointment. Another clock time or therapist INSIDE parentheses = linked replacement, **not an independent second appointment**. Parentheses for weekdays/notes are not automatically replacement; never infer missing therapist, time or calendar date from file name.
- Analysis of 271 previously indexed visible-sheet cells yielded 271 individual **HOLD** staging rows (no patient names or raw notes copied); 7 with an explicit time inside parentheses (candidate linked replacements); 9 cells with more than one explicit time *outside* parentheses (ambiguous, HOLD). Values counted are **cell-level signals, not validated sessions or replacement KPIs**.
- Staging artifact created in current conversation sandbox only: `DIOMIL_MANUAL_SCHEDULE_STAGING_REVIEW_20261110.xlsx`, three sheets `ΣΥΝΟΨΗ`, `ΠΡΟΣ ΕΠΑΛΗΘΕΥΣΗ`, `ΚΑΝΟΝΕΣ & ΟΡΙΑ`, 271 reviewer IDs and 271 `HOLD - NO IMPORT`. Human-facing methodological report `DIOMIL_MANUAL_SCHEDULE_STAGING_READ_ONLY_20261110.md`. The source hash independently rechecked; staging .xlsx ZIP CRC clean, exactly three worksheets.
- Time format normalized into 24h text for **first unparenthesized explicit time only**; column for parenthesized time kept as provisional linked replacement. First therapist not yet parsed/validated against authoritative therapist registry, dates not yet resolved, row-to-patient linkage only points back to cell A[row] (no copied patient names). No appointment/slot/attendance writeback to Diomil occurred.
- Next admission dependency: validated therapist aliases and IDs, weekday versus dated-program semantics, distinguish therapist parenthesis vs day/notes, strict clock/slot conflicts and explicit operator review of ambiguous cases. Do not claim pilot-ready on this staging preview. Keep frozen synthetic H10 CC and real manual source fully separate.


### 2026-11-11 — Operator prioritizes fast pilot delivery: scope freeze and one acceptance pass
User requests to finish promptly; incremental cell-by-cell parser extensions must cease unless they prevent core safe pilot operation. The real manual workbook is a noisy *weekly working reference*, not a trustworthy dated clinical ledger; do not invent calendar dates or infer live attendance from its filename.
1. Freeze current C and E/H parsers and overlap audit at **35 Windows tests passed**. Preserve agreed source semantics. Reuse existing patient absence workflow, no new categories.
2. Next delivery is one combined **read-only candidate import manifest** (not live writeback): source cell, row/patient linkage, specialty, explicit day+time, initial vs replacement therapist where verified, confidence and unresolved reason; for each distinct status provide counts and source examples without exposing patient names in public reporting. Do not force missing weekdays into E/H or missing times into G.
3. Date binding is an explicit admission gate. Where operator has not established the reference/effective week, treat source items as *weekly-pattern proposals* rather than appointments on specific dates. Ask only the minimal targeted question if real calendar placement is actually required; otherwise prepare preview autonomously.
4. Check one end-to-end preview against registry aliases, timeslot/capacity, same-patient overlaps and user-visible correction workflow. Hold ambiguous cells for operator, never silently approve them; never overwrite frozen BZ/CC or original workbook.
5. Promote to an isolated parallel pilot only after explicit source-to-target mapping and native Excel evidence; do not claim that 35 green unit tests mean the real dated clinical schedule is ready. Avoid a new round of cosmetic changes or unrelated feature development.


### 2026-11-11 — Operator-specified census = 100 patients, targeted review only
The operator explicitly states **100 patients** in the manual source and authorizes asking *for each specific unrecognized patient row* rather than extending speculative parsing. Treat 100 as the **expected patient census for reconciliation**, not as inferred number of clinical sessions. Read-only check of original visible `ΠΡΟΓΡΑΜΜΑ` shows 107 nonempty column-B cells spanning B2..B111, including non-patient/header/annotation candidates; these **are NOT 107 patients**. Do not equate populated B cells with patient records. Filter titles/notes and row semantics, then reconcile to 100 *distinct patient rows*; report only ambiguous **exact cell references** (B## and adjacent A##/services) for user clarification, without exposing patient names in chat. Keep source untouched and do not create/import patient records until ambiguities and identity mapping are reviewed. Avoid prolonged speculative investigations; user prefers completion with limited targeted questions.


### 2026-11-11 — Consolidated real manual schedule review artifact PRODUCED (not Diomil import)
Responding to operator request to stop incremental steps, built **one finished editable review workbook** `DIOMIL_ΕΝΙΑΙΟΣ_ΕΛΕΓΧΟΣ_100_ΑΣΘΕΝΩΝ_20261111.xlsx`, available in conversation, **3 sheets**: `ΑΣΘΕΝΕΙΣ_ΕΛΕΓΧΟΣ` (105 candidate patient rows, original B identity and A–J content, editable confirmation/comments); `ΘΕΡΑΠΕΙΕΣ_ΚΕΛΙΑ` (324 nonempty service cells for those rows with conservative C and earlier E/H proposals, original raw source text, editable review/comments); `ΣΥΜΦΩΝΙΑ_100` (census reconciliation, source checksum, no-import warning). Excel ZIP validated, source SHA256 unchanged `765264ec8eae95a31ccc57629dc31b1d61f59ea5c992221affa91fcf6a8d2567`.
- 105 B-column candidate patient rows counted on visible sheet with B4..B111 excluding clinic header B58 and blanks B99/B103. Exactly 5 of 105 include older source annotations: **B17, B20, B64, B72, B86**; **do not infer** these five automatically must be excluded or represent five inactive patients. Reconcile 105 to operator's expected 100 with one targeted user determination, never silently discard clinical identities.
- Distinguish source content from validated import. Patient names are present in the **private editable review workbook** so that operator can make practical corrections, but not repeated in chat/reports. Existing patient absence/status workflows must not be reconstructed from B markers.
- Earlier C parser proposal CSV and earlier E/H CSV are partial and are included as **proposals only**, not calendar-date validated. No new date/therapist-ID mapping or timeslot allocation, no writeback to production Diomil, no changes to synthetic CC. Do not call this pilot ready.
- This closes preliminary read-only consolidation. Next targeted decision is membership of the five flagged source rows in current 100; then only a single controlled pilot import with source-to-target mapping and native safety gates, not another broad manual parser project.
