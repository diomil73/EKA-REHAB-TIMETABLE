# DIOMIL — Συγκεντρωτική αναφορά παραδόσεων και εκκρεμοτήτων

## 2026-11-16 — Executable private real-source DEMO preflight

**New executable interface:** `Python/tools/prepare_real_data_practice_demo.py` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Default operation is **read-only full real-source validation** for an authorized current A:V manual XLSX, confirmed review XLSX, entirely empty application-ready VBA XLSM and independently approved private `clinical_approvals.json`. Uses the strict production source alignment, all-patient approved ID check, original weekday/time/provider reconciliation, unresolved F/U/V stop gates, existing native shell checks and SHA-256 stability of the clinical-approval JSON. Every input must remain off Git checkout. No real patient name, ID, room, doctor, MRN, source path or clinical exception content is emitted to stdout/stderr on handled errors. `--build --output <PRIVATE_*_DEMO_*.xlsm>` is an *explicit* Windows + native Excel/COM, exclusive-new-file second step and repeats the same approval gates. **No real source file has been run; no real-data XLSM exists, and end-to-end practical acceptance is still pending.** Previous synthetic/H8 workbook results do not count as the operator's comparison. First actual action is an authorized private preflight with current manual and approved clinical review, not a random synthetic seed. Native failure or changed approvals after a native operation requires freezing the candidate, not treating it as accepted. See the dedicated operator CLI runbook for exact commands/JSON schema.

**Current blocker:** F/U/V and independently sourced outpatient bookings remain unavailable to the current *inpatient-only* A:V admission bridge. If occupied, the complete real-data source fails closed, never silently drops treatments. Do not alter frozen H2-H10 evidence. Update all living handoff/status/test and delivery documents again with actual source-specific counts and results **only after** the first private native execution, without publishing PII.

## ΝΕΟ ΚΡΙΤΗΡΙΟ ΑΠΟΔΟΧΗΣ 2026-11-16: REAL-DATA DEMO

**Νέο authoritative addendum (χωρίς αναθεώρηση παλαιών H2–H10 proofs):** Η πραγματική επιχειρησιακή αξιολόγηση πρέπει να γίνει με **επικαιροποιημένο πραγματικό χειροκίνητο πρόγραμμα** και 1:1 σύγκριση ασθενών/ειδικοτήτων/ημερών/ωρών/θεραπευτών. Οποιοδήποτε παλιό synthetic-only demo milestone είναι τεχνικό regression fixture και **όχι αποδεκτό παραδοτέο**.

**Confirmed engineering evidence:** `9fcc6108` και GitHub Actions `43769366651`, 1.960 GREEN Windows + Linux, synthetic gates GREEN. **Not delivered:** πραγματικό native XLSM, current-source read-only alignment/approved private identities/clinical facts, πλήρης F/U/V, independent COM readback/user acceptance, in-session active MASTER update, admissions/absence/replacement/staff/student/save-restore integration. Δεν υπάρχουν προσωπικά δεδομένα στο repo. Επόμενο acceptance register και full operator runbook: `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`; παλαιές specialist checkpoints frozen.

Ημερομηνία απογραφής: 2026-11-15  
Repository: diomil73/EKA-REHAB-TIMETABLE  
Branch απογραφής: work/afternoon-medical-scope-20261115  
Τύπος: evidence-based project delivery register, όχι δήλωση production release.

---

## 1. Γιατί δημιουργήθηκε αυτή η αναφορά

Πριν από αυτή την απογραφή υπήρχαν δύο χρήσιμα αλλά διαφορετικά έγγραφα:

- PROJECT_HANDOFF_EVIDENCE_2026-11-03.md: ιστορική παράδοση κυρίως για H2–H7.
- INTEGRATED_TEST_BASELINE_MATRIX_20261115.md: αποτύπωση αποδεκτών native artifacts, scope και μεγάλων κενών για τη δημιουργία ενός πλήρους synthetic integration workbook.

Κανένα από τα δύο δεν ήταν πλήρης, ενιαία αναφορά παραδόσεων όλων των workstreams του Diomil. Το παρόν έγγραφο καλύπτει αυτό το κενό. Δεν αντικαθιστά τις επιμέρους αναφορές, τα runbooks ή τα πρωτογενή logs. Συγκεντρώνει την κατάσταση και παραπέμπει στα πρωτογενή στοιχεία.

### Απογραφή τεκμηρίωσης

Στο Documentation εντοπίστηκαν 149 αρχεία. Από αυτά, με βάση τα ονόματά τους:

- 35 αρχεία με CHECKPOINT στο όνομα,
- 46 αρχεία της σειράς RECOVERY_,
- 8 αρχεία handoff / resume / status / onboarding,
- 5 native runbooks,
- 5 H10-specific guides / runbooks.

Οι κατηγορίες επικαλύπτονται και δεν αθροίζονται. Τα 149 αρχεία περιλαμβάνουν επίσης συμβόλαια, αποφάσεις, προδιαγραφές, σχεδιασμούς, audits και οδηγίες. Δεν είναι 149 ολοκληρωμένα παραδοτέα.

Η απογραφή εξέτασε τα βασικά handoff/status έγγραφα, την acceptance matrix, τις τελικές H2–H10 αναφορές, τις βασικές αναφορές Recovery, την πραγματική πηγή/εισαγωγή και το τρέχον afternoon scope. Τα ενδιάμεσα failed checkpoints διατηρούνται ως ιστορικό και δεν υπερισχύουν των νεότερων τελικών αποδείξεων.

---

## 2. Πώς διαβάζονται οι καταστάσεις

| Κατάσταση | Ακριβής σημασία |
|---|---|
| NATIVE ACCEPTED / FROZEN | Ο χρήστης επιβεβαίωσε συγκεκριμένη λειτουργία στο συγκεκριμένο Excel artifact. Το αρχείο είναι αποδεικτικό, όχι writable βάση. |
| NATIVE INSTALL / READ-ONLY AUDIT GREEN | Εγκατάσταση και persisted readback πέρασαν. Δεν συνεπάγεται απαραίτητα χειροκίνητη αποδοχή UI. |
| SOURCE GREEN ONLY | Python/source tests πέρασαν. Δεν αποδεικνύει native Excel ή λειτουργική αποδοχή. |
| PARTIAL ACCEPTANCE | Έγινε αποδεκτό συγκεκριμένο υποσύνολο, όχι ολόκληρη η απαίτηση ή το workstream. |
| DESIGN / NOT IMPLEMENTED | Υπάρχει απαίτηση ή σχεδιασμός, χωρίς αποδεκτή υλοποίηση. |
| BLOCKED | Λείπει υποχρεωτικό source, provenance, επιχειρησιακή απόφαση ή native proof. |
| FAILED / FROZEN | Η συγκεκριμένη προσπάθεια απέτυχε και διατηρείται ως διαγνωστικό στοιχείο. Δεν επαναχρησιμοποιείται. |

Ένα GREEN αφορά πάντα το συγκεκριμένο scope της αντίστοιχης αναφοράς. Δεν σημαίνει ότι ολόκληρο το Diomil είναι έτοιμο.

---

## 3. Συνολική εικόνα

### Τι έχει ουσιαστικά παραδοθεί

1. Ο βασικός μηχανισμός ημερήσιας ανακατασκευής και οι σημαντικές διαδρομές απουσίας / αντικατάστασης / επαναφοράς έχουν ισχυρή source και native τεκμηρίωση.
2. Το κλείσιμο ημέρας, η διατήρηση ιστορικού και η versioned διόρθωση κλεισμένης ημέρας έχουν native proofs.
3. Τα ιστορικά στατιστικά θεραπευτών και ασθενών έχουν λειτουργικό backend και αποδεκτά συγκεκριμένα Excel UI scopes.
4. Το H8 dashboard έχει native και χειροκίνητη αποδοχή για το καταγεγραμμένο dashboard scope.
5. Έχει αποδειχθεί σε synthetic αντίγραφα η native αποθήκευση/reopen και η μεταφορά 179 σταθερών recurring AssignmentIDs.
6. Το outpatient CREATE και βασικά registration / preview backends υπάρχουν και έχουν επιμέρους τεκμηρίωση.

### Τι δεν έχει παραδοθεί ως ολοκληρωμένο προϊόν

1. Πλήρης καταγραφή πραγματικών παρουσιών/απουσιών θεραπευτών με αξιόπιστο ημερολόγιο εργασίμων και ιστορικό.
2. Μόνιμη αλλαγή recurring θεραπείας με native event append, transaction, rollback και ολοκληρωμένο UI.
3. Πραγματικός importer / native admission του ενημερωμένου προγράμματος ασθενών.
4. Πλήρης, αποδεδειγμένα synthetic, ενιαία δοκιμαστική έκδοση ολόκληρης της εφαρμογής.
5. Ολοκληρωμένη πολιτική προστασίας όλων των φύλλων και των επιτρεπόμενων περιοχών επεξεργασίας.
6. Ο global optimizer για ταυτόχρονες απουσίες πολλών θεραπευτών.
7. Η 15ήμερη αξιολόγηση / πρόοδος ασθενούς.
8. Native integration και save path για το νέο afternoon physiotherapy / medical deferral scope.
9. Πλήρης αποδοχή των τελικών H10 charts και του ανεξάρτητου management XLSX.

---

## 4. Αναλυτικό μητρώο παραδόσεων

### 4.1 Ημερήσια λειτουργία, ανακατασκευή και αντικαταστάσεις

**Κατάσταση: CORE MECHANISMS NATIVE-PROVEN, με ανοιχτό global optimizer.**

Η ημερήσια αλήθεια βασίζεται στη διαδρομή:

recurring/base schedule + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS → DailySessionState → operational views/statistics.

Οι απουσίες, ακυρώσεις, no-shows, αναστολές και αντικαταστάσεις είναι overlays και δεν πρέπει να καταστρέφουν το recurring πρόγραμμα.

Στις αναφορές Recovery τεκμηριώνονται:
- restart-safe authoritative reconstruction,
- replacement lifecycle V87–V92,
- authoritative DAILY_REPLACEMENT_FACTS,
- per-row / selected / restore-all restore backend,
- validation πριν από mutation, semantic readback και rollback,
- συνέπεια MASTER_SCHEDULE / THERAPIST_DAILY / daily statistics,
- in-session mutation αντί για παλιό close/reopen preview-sync μοντέλο.

**Native patient-absence Apply / Restore X:** επιβεβαιώθηκε μία συγκεκριμένη all-day απουσία ασθενούς, Apply μία φορά και Restore μία φορά. Επαληθεύτηκαν:
- αφαίρεση και επαναφορά από MASTER και THERAPIST_DAILY,
- επαναφορά του MASTER_SCHEDULE!D8 στην εσωτερική PATIENT_PLANNER formula,
- επαναφορά του MASTER_SCHEDULE!M8 σε παρών,
- DAILY_CHANGE_LEDGER ACTIVE → RESTORED,
- συμπλήρωση RestoredAt,
- εκκαθάριση της αντίστοιχης DAILY_INPUT γραμμής.

Αυτό δεν αποδεικνύει κάθε συνδυασμό απουσιών ή κάθε μελλοντική recurring αλλαγή.

**Ανοιχτό:** η ποιότητα ανάθεσης όταν υπάρχουν πολλές ταυτόχρονες απουσίες θεραπευτών. Το single-patient candidate ranking δεν είναι ο τελικός global multi-absence optimizer. Οι κανόνες scarcity, adjacency, capacity, workload balance και explicit exceptions πρέπει να διατηρηθούν.

Πηγές:
- RECOVERY_RUNTIME_PAYLOAD_CAPTURE_GATE_2026-11-01.md
- RECOVERY_J_TO_M_HANDOFF_2026-10-31.md
- RECOVERY_GATE_67_GREEN_2026-10-31.md
- RECOVERY_CANONICAL_SOURCE_GREEN_GATE_2026-11-01.md
- RECOVERY_DAILY_CHANGE_RESTORE_2026-10-31.md
- RECOVERY_REPLACEMENT_OPERATOR_RESTORE_2026-10-31.md
- CURRENT_HANDOFF.md

### 4.2 H2 — Κλείσιμο ημέρας και durable history

**NATIVE ACCEPTED / FROZEN**

Το Z είναι το πρώτο operator-confirmed native close-day proof της συγκεκριμένης recovery γραμμής:
- fresh workbook build,
- read-only preflight,
- επιτυχές CloseOperationalDay μία φορά,
- επιτυχής post-close read-only inspection.

Το ακριβές inspector output δεν είχε επικολληθεί στην αναφορά, επομένως δεν αποδίδονται επιπλέον counts ή IDs που δεν τεκμηριώθηκαν.

Artifact: UNIFIED_CLOSE_DAY_HISTORY_20261101_Z.xlsm.

Πηγή: CLOSE_DAY_NATIVE_Z_PROOF_2026-11-01.md.

### 4.3 H3 — Διόρθωση κλεισμένης ημέρας / version chain

**NATIVE ACCEPTED / FROZEN**

Το AB απέδειξε:
- source gate 35 passed in 3.23s,
- read-only preparation και hash-bound apply,
- native atomic persistence,
- version 1 → version 2,
- v1 preserved ως SUPERSEDED,
- v2 CURRENT_FINAL,
- 16 therapist/day history rows,
- linked amendment STARTED / FINALIZED events,
- independent post-save outcome inspection.

Αυτό αποδεικνύει το durable backend. Η user-facing correction action / navigation ήταν ξεχωριστό UX θέμα, το οποίο εξελίχθηκε αργότερα μέσα στο H6 contextual Save Day. Το H3 proof από μόνο του δεν αποδεικνύει όλη τη μεταγενέστερη εφαρμογή.

Artifact: UNIFIED_DAY_AMENDMENT_20261101_AB.xlsm.

Πηγή: DAY_AMENDMENT_NATIVE_AB_PROOF_2026-11-01.md.

### 4.4 H4 — Ιστορικές αναφορές θεραπευτών

**NATIVE ACCEPTED / FROZEN**

Backend και UI αποδείχθηκαν για:
- Ημερήσια,
- Εβδομαδιαία,
- Μηνιαία,
- Συνολικά.

Η αναφορά διαβάζει finalized history και μετρά μόνο CURRENT_FINAL. Τα SUPERSEDED snapshots παραμένουν στο audit history, αλλά δεν διπλομετρώνται.

Μετρικές:
- planned / actual occupied timeslots,
- planned / active sessions,
- replacements received,
- released load από cancellations/no-shows,
- treatment suspensions,
- planned / actual capacity utilization.

Κρίσιμη σημασιολογική διάκριση: occupied clock slots ≠ αριθμός ασθενών/συνεδριών. Shared-slot appointments μπορούν νόμιμα να έχουν περισσότερες συνεδρίες από occupied timeslots.

Το AM χρησιμοποιήθηκε read-only για report proof. Το AQ είναι το τελικό native UI artifact όπου ο χρήστης επιβεβαίωσε εναλλαγή όλων των modes και παραμονή των κουμπιών.

Artifacts:
- UNIFIED_DAY_SAVE_UI_20261102_AM.xlsm — read-only report proof.
- UNIFIED_THERAPIST_HISTORY_UI_20261102_AQ.xlsm — native UI proof.

Πηγές:
- H4_THERAPIST_HISTORY_REPORTS_CHECKPOINT_2026-11-02.md
- H4_THERAPIST_HISTORY_REPORTS_GREEN_2026-11-02.md
- CHECKPOINT_H4_THERAPIST_HISTORY_2026-11-03.md

### 4.5 H5 — Ιστορικό συνεδριών ασθενών

**NATIVE ACCEPTED / FROZEN**

AX:
- native first close για 2026-10-30,
- 80 patient-session rows,
- 16 therapist history rows,
- v1 CURRENT_FINAL,
- finalized event και τεχνικά history sheets veryHidden.

AY:
- native amendment v1 → v2,
- previous patient/session rows 80,
- current patient/session rows 80,
- operational CURRENT_FINAL rows 80,
- v1 preserved, v2 current.

Τα AX και AY είναι ξεχωριστά immutable proofs.

Πηγή: CHECKPOINT_H5_PATIENT_HISTORY_2026-11-03.md.

### 4.6 H6 — Contextual Save Day UX

**NATIVE ACCEPTED / FROZEN**

Το AZ απέδειξε λειτουργικό contextual Save Day flow με v1 → v2. Το BB επιβεβαίωσε την τελική διάταξη:
- Save & Exit παραμένει app-shell action,
- το Save Day βρίσκεται στο MASTER,
- δεν υπάρχουν παλιά close/amend/save controls στο DAILY_INPUT,
- ο dispatcher υπάρχει,
- ο χρήστης επιβεβαίωσε οπτικά τη διάταξη.

Σημείωση scope: στο BB δεν εκτελέστηκε νέο save action. Η λειτουργική v1→v2 απόδειξη προέρχεται από το AZ, ενώ το BB αποδεικνύει την τελική διάταξη.

Artifacts: UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm και ..._BB.xlsm.

Πηγές:
- CHECKPOINT_H6_CLOSED_H7_STARTED_2026-11-03.md
- CHECKPOINT_H6_MASTER_ONLY_SAVE_2026-11-03.md
- CHECKPOINT_H6_UNIFIED_DAY_SAVE_UX_2026-11-03.md
- PROJECT_HANDOFF_EVIDENCE_2026-11-03.md

### 4.7 H7 — Ενοποίηση ιστορικών αναφορών και toolbar

**PARTIAL ACCEPTANCE, LATER REFINED BY H8**

Η σειρά BC / BD / BE περιλαμβάνει χρήσιμα διαγνωστικά:
- BC: εγκατάσταση και αρχική οπτική παρατήρηση, όχι πλήρης integrated acceptance.
- BD: persisted structure audit green, αλλά υπήρξαν intermittent Overall report error και ερωτήματα patient coverage.
- Η σύγκριση PATIENTS registry 93 έναντι CURRENT_FINAL 77 distinct patients / 80 sessions απέδειξε ότι 16 registered patients δεν είχαν finalized session στο συγκεκριμένο test date. Δεν επιτρέπεται να δημιουργούνται πλασματικές συνεδρίες.
- Οι IDs 18 και 22 υπήρχαν στα report rows· η απουσία τους από την οπτική αντίληψη δεν αποδείχθηκε ως data loss.
- BI είχε installer και protection audit green, αλλά ο χρήστης απέρριψε το analytics UX.

Το BF είναι το λειτουργικά αποδεκτό integrated H7 smoke:
- 26 focused source tests passed in 4.87s,
- integrated toolbar,
- retained RefreshAll,
- patient και therapist report entry points,
- report modes,
- previous-screen navigation,
- independent H7 integrated read-only audit,
- χειροκίνητη επιβεβαίωση χρήστη «όλα γίνανε σωστά».

Το BF δεν σημαίνει ότι ολοκληρώθηκε η μεταγενέστερη απαίτηση για πλήρη analytics UX. Αυτή οδήγησε στο H8.

Artifact: UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm.

Πηγές:
- CHECKPOINT_H7_BF_MANUAL_ACCEPTANCE_AND_PROTECTION_NEXT_2026-11-04.md
- CHECKPOINT_H7_BD_INTEGRATED_AUDIT_GREEN_2026-11-03.md
- CHECKPOINT_H7_REGISTRY_93_VS_HISTORY_77_2026-11-04.md
- CHECKPOINT_H7_REFRESHALL_RETAINED_AND_TOOLBAR_MIGRATION_2026-11-04.md
- H8_STATISTICS_DASHBOARD_OPERATOR_FEEDBACK_2026-11-07.md

### 4.8 H8 — Patient / therapist dashboards

**NATIVE ACCEPTED / FROZEN FOR RECORDED SCOPE**

Το BL έγινε αποδεκτό μετά από:
- 74 passed in 4.55s source gate,
- native install και persisted VBA checks,
- independent protection / integrated UI audits,
- χειροκίνητη επιβεβαίωση «όλα καλά».

Αποδεκτό scope:
- patient overview και selected-patient charts,
- Προγραμματισμένες / Πραγματοποιημένες / Απουσίες,
- active period-button styling,
- sticky patient chart κατά το scroll,
- τρία therapist dashboard panels με σαφή όρια διαθεσιμότητας,
- report navigation / protection,
- ένα τελικό RefreshAll success notification.

Το τρίτο therapist panel δεν είναι πραγματικό quantitative attendance chart. Παραμένει ρητά «Μη διαθέσιμο ιστορικό απουσιών / εργασίμων» μέχρι να υπάρξει authoritative attendance source.

Το BO synthetic history proof επιβεβαίωσε:
- 25 service days,
- 522 patient-session rows,
- 150 therapist-day rows,
- 150 synthetic attendance facts,
- 25 finalization events,
- cross-sheet consistency και unchanged source.

Τα BM και BN απέτυχαν κατά το synthetic native seeding και διατηρούνται ως frozen diagnostics. Η τελική επιτυχία BO δεν μετατρέπει τα synthetic attendance facts σε πραγματικές παρουσίες.

Artifacts:
- UNIFIED_PATIENT_HISTORY_H8_20261107_BL.xlsm — accepted UI.
- UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm — synthetic history evidence.

Πηγές:
- H8_STATISTICS_DASHBOARD_OPERATOR_FEEDBACK_2026-11-07.md
- H8_SYNTHETIC_HISTORY_DEMO_PLAN.md
- PROJECT_STATUS.md, H8 appendices.

### 4.9 H9 — Ιστορικό παρουσιών και εργασίμων θεραπευτών

**SOURCE / READ-ONLY MODEL ONLY**

Το H9 εντόπισε πραγματικό κενό στο υπάρχον schema: therapist/day workload δεν αποδεικνύει αν ο θεραπευτής ήταν υποχρεωμένος να εργαστεί μια συγκεκριμένη ημέρα. Η απουσία συνεδριών δεν ισοδυναμεί με απουσία θεραπευτή.

Υλοποιήθηκε pure domain contract για PRESENT / ABSENT / NONWORKING, με provenance, reason και explicit expected-working-date calendar. Ο operator επιβεβαίωσε 7 passed in 0.98s για το pure-domain contract.

Προστέθηκε read-only store reader με schema:
SnapshotId, ServiceDate, TherapistId, Outcome, Source, ReasonCode.

Δεν έχει αποδειχθεί:
- live finalized attendance writer,
- αυτόματη παραγωγή roster/calendar,
- holiday/leave policy,
- σύνδεση με day close / amendment writer,
- native Excel install ή quantitative chart.

Το H8 demo attendance sheet είναι illustrative synthetic evidence μόνο.

Σημείωση ελέγχου αντιφατικών checkpoints: το H8 plan αναφέρει 13 passed in 8.33s για separate H9 model/read-only loader. Η μεταγενέστερη H9 contract αναφέρει 7 passed για το pure contract και μετά προσθέτει store tests που δηλώνονται ως pending. Μέχρι να βρεθεί το ακριβές νεότερο combined test output, δεν χαρακτηρίζεται ο store reader πλήρως verified.

Πηγή: H9_THERAPIST_ATTENDANCE_HISTORY_CONTRACT.md.

### 4.10 H10 — Therapist dashboard, management view και synthetic proofs

**PARTIAL ACCEPTANCE / NOT FULLY ACCEPTED**

Υπάρχουν τρεις διαφορετικές κατηγορίες H10 evidence:

**Α. Therapist UI / charts**
- CC native install, persisted VBA, history cross-sheet verification και strict capacity audit πέρασαν.
- 25 synthetic days, 522 patient-session facts, 150 therapist/day facts, 150 synthetic attendance facts, 25 finalization events.
- 150 therapist/day rows audited. Τα delivered sessions μπορούν να υπερβαίνουν το initial plan λόγω replacements, αλλά τα occupied clock slots πρέπει να αξιολογούνται χωριστά.
- Ο χρήστης έστειλε screenshots και αποδέχθηκε οπτικά δύο συγκεκριμένες οθόνες: το Συνολικά therapist report και το internal MANAGEMENT_SUMMARY. Αυτό δεν ισοδυναμεί με πλήρη functional acceptance όλων των έξι periods, όλων των κουμπιών ή κάθε chart path.
- Το δυναμικό therapist profile για 20+ θεραπευτές παραμένει design-only.

**Β. Management-only standalone XLSX**
- Υπάρχει ανεξάρτητο export path και σχεδιασμός για τέσσερα macro-free sheets με native shape hyperlinks, hidden headings/gridlines/tabs και KPI tiles.
- Το export είναι ξεχωριστό artifact από το operational XLSM.
- Η πλήρης privacy/readback/visual approval για ασφαλή κοινοποίηση δεν πρέπει να συναχθεί από screenshots ή source markers μόνο.

**Γ. H10 synthetic native proof chain**
- CC native read-only probe A: αποδεικνύει isolated native Excel open/close read-only, όχι writeback.
- Native Save/reopen A: αποδεικνύει προσθήκη του DIOMIL_NATIVE_PROBE marker, Save/close/reopen, VBA preservation και unchanged synthetic history.
- Recurring AssignmentID migration A: αποδεικνύει native persistence/readback 179 immutable AssignmentIDs.
- Δεν αποδεικνύεται permanent recurring event append, authoritative transaction, rollback ή production migration.

Το H10 therapist absence/working-day chart παραμένει unavailable χωρίς genuine finalized attendance facts. Δεν εμφανίζουμε 0 απουσίες από ελλιπή δεδομένα.

Πηγές:
- H10_THERAPIST_CHARTS_USER_REQUIREMENTS_20261108.md
- H10_NEXT_SESSION_READ_FIRST_20261108.md
- H10_BV_NATIVE_RUNBOOK_20261108.md
- H10_BW_NATIVE_RUNBOOK_20261108.md
- NATIVE_CC_READONLY_PROBE_20261112.md
- NATIVE_CC_SYNTHETIC_SAVE_REOPEN_20261112.md
- NATIVE_SYNTHETIC_RECURRING_ID_MIGRATION_20261112.md
- AUTHORITATIVE_RESUME_20261112.md

### 4.11 Outpatient recurring schedule

**Κατάσταση: CREATE NATIVE GREEN, EDIT / CONFLICT / WORKLOAD DETAILS PARTIAL.**

Η Implementation Map καταγράφει:
- outpatient CREATE native green,
- application-style schedule manager native-visible and usable,
- stable same-row edit semantics implemented,
- duplicate edit projection bug hardened και latest fresh native smoke appears correct,
- workload-strip UI implemented in source αλλά awaiting focused/native verification,
- separate outpatient conflict native gate still pending.

Το OUTPATIENT_SCHEDULE είναι hidden/protected authoritative storage, όχι end-user worksheet. Το public edit flow πρέπει να χρησιμοποιεί business-facing manager και όχι row numbers.

Πηγή: IMPLEMENTATION_MAP.md, ενότητα Outpatient recurring scheduling.

### 4.12 Registries και entity registration

**Κατάσταση: VALIDATED BACKENDS / PREVIEW WRITEBACK, ΟΧΙ ΠΛΗΡΗΣ ΕΝΙΑΙΑ NATIVE REGISTRATION ACCEPTANCE.**

Υπάρχουν:
- patient registry validation με stable PatientID,
- therapist registry validation από SETTINGS,
- dedicated STUDENTS A:H schema,
- preview copy + native Excel writeback/readback για τις αντίστοιχες entity flows,
- registration orchestrator και form specifications για patient / therapist / student.

Το registration menu contract παραμένει UI-facing specification. Η formal registration αναφορά διευκρινίζει ότι το orchestration είναι backend plumbing και δεν προσθέτει από μόνο του νέα VBA UserForms στο binary XLSM ή automatic patient assignment.

Πηγές:
- FORMAL_ENTITY_REGISTRATION.md
- REGISTRATION_MENU_CONTRACT.md
- IMPLEMENTATION_MAP.md
- student registry / transition specialist docs.

### 4.13 Permanent recurring therapy changes

**SOURCE PLANNING + STABLE ID NATIVE PROOF; EVENT WRITE NOT PROVEN**

Έχουν υλοποιηθεί:
- effective-dated recurring change domain projection,
- read-only workbook preview με SHA/stale-row checks,
- immutable AssignmentID binding,
- synthetic-only migration plan,
- native migration 179 AssignmentIDs,
- read-only RECURRING_CHANGE_EVENTS parser,
- source-SHA-bound in-memory append plan.

Έχει καταγραφεί CI run 43314844755 με 1,652 full Python tests σε Windows και Linux για την αντίστοιχη source revision.

Δεν έχει αποδειχθεί:
- native append event σε workbook,
- authoritative commit transaction,
- compensating/reversal event rollback,
- full operational VBA Apply UI,
- real patient migration.

Πηγή: AUTHORITATIVE_RESUME_20261112.md και IMPLEMENTATION_MAP.md.

### 4.14 Real-source trial / admission

**BLOCKED — NO ACCEPTED REAL-DATA IMPORTER**

Υπάρχουν read-only tools:
- trial_admission_gate,
- preflight_trial_excel_sources,
- audit_pilot_source_readiness,
- trial_import_plan,
- trial_weekly_projection.

Το trial plan κάνει validation, hash binding και projection σε υπάρχον schema. Δεν γράφει XLSM και δεν είναι importer.

Η αναφορά 2026-11-13 απαιτεί:
- reviewed private source,
- confirmed PatientID mapping,
- μηδενικά unresolved holds και overlaps ή ξεχωριστή ρητά εγκεκριμένη quarantine policy,
- exact shared-slot membership approval,
- νέο κατάλληλο private candidate,
- native writer, independent readback και privacy proof.

Οι παλαιές μετρήσεις 61 held cells / 11 candidate overlaps είναι historical staging observations, όχι σημερινό validated census. Η real-data demo δεν έχει ολοκληρωθεί ή γίνει native accepted.

Καμία πραγματική κλινική λίστα, patient row ή private workbook δεν πρέπει να ανέβει στο public repository ή σε public logs.

Πηγή: REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md και AUTHORITATIVE_RESUME_20261112.md.

### 4.15 Afternoon physiotherapy / pool / medical deferral

**SOURCE GREEN ONLY — NATIVE PROOF BLOCKED BY PROVENANCE**

Η πολιτική ορίζει:
- pool ως πρωινή θεραπεία,
- afternoon programme ως ξεχωριστή physiotherapy-only namespace,
- PM fields αρχικά κενά,
- ανεξάρτητα morning/afternoon conflict namespaces,
- προσωρινή discharge/transfer ως patient-wide state,
- medical deferral με σαφή scope: whole day, time window ή exact session,
- καμία υπόθεση duration από start time,
- preview και ρητή επιβεβαίωση των επηρεαζόμενων sessions.

Στο active branch υπάρχουν pure policy, draft, read-only source schema και read-only patient form panel. Η πιο πρόσφατη επιβεβαιωμένη focused Windows suite ήταν 49 passed in 9.22s. Η πλήρης regression μετά από το VBA fixture fix ήταν 1,808 passed, 1 skipped, 0 failed σε 30.54s· μεταγενέστερη πλήρης regression μετά τον integrated structure auditor ήταν 1,813 passed, 1 skipped, 26 warnings.

Δεν υπάρχει ακόμη:
- native Excel save path για afternoon slots,
- native acceptance του panel,
- operational medical deferral UI / transaction,
- rollback / amendment integration,
- πλήρες synthetic integrated workbook για ασφαλές native smoke.

Το native installer παραμένει gated επειδή δεν έχει ταυτοποιηθεί πλήρως synthetic integrated source με τα σωστά PatientProgramRuntime / frmPatientScheduling components. Το populated canonical workbook εξαιρείται.

Πηγή: AFTERNOON_PHYSIOTHERAPY_MEDICAL_DEFERRAL_POLICY_20261115.md.

### 4.16 Worksheet protection

**Κατάσταση: PARTIAL PROTECTION, BROAD RELEASE GATE OPEN.**

Το BF read-only protection audit κατέγραψε:
- 23 worksheets,
- 18 unprotected,
- 5 protected,
- ProtectStructure=False,
- ProtectWindows=False,
- τα UsedRange cells να εμφανίζονται Locked=True δεν εμποδίζει edits όταν ProtectContents=False.

Οι δύο report sheets προστατεύτηκαν στο H7/H8 chain και τα τεχνικά history sheets παραμένουν veryHidden/protected. Αυτό δεν είναι συνολική προστασία όλων των worksheets.

Η τελική προστασία πρέπει να βασιστεί σε πλήρη χαρτογράφηση manual edit zones και όλων των VBA/Python write paths. Blanket locking χωρίς αυτή τη χαρτογράφηση μπορεί να σπάσει νόμιμες φόρμες και COM writers.

Πηγές:
- BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md
- WORKSHEET_PROTECTION_WRITE_PATH_MAP_2026-11-04.md
- CHECKPOINT_H7_BF_MANUAL_ACCEPTANCE_AND_PROTECTION_NEXT_2026-11-04.md

### 4.17 Patient 15-day progress / evaluation

**DESIGN ONLY**

Το STATISTICS_HISTORY_DESIGN.md περιγράφει:
- manual evaluation περίπου ανά 15 ημέρες,
- append-only milestones,
- question/goal, Από, Έως, outcome,
- πράσινο = πρόοδος, γκρι = ουδέτερο, κόκκινο = επιδείνωση,
- PatientID-linked trend που δεν αλλάζει όταν αλλάζει θεραπευτής/ώρα/κελί.

Δεν υπάρχει native-accepted workflow στο evidence set που εξετάστηκε. Δεν πρέπει να δηλώνεται ως υλοποιημένο.

### 4.18 Whole-workbook synthetic integrated test fixture

**NOT CERTIFIED**

Η πλήρης synthetic integration workbook παραμένει μη πιστοποιημένη.

Το tracked Rehab_Center_System_v27_1.xlsm είναι populated clinical baseline και εξαιρείται από native write tests. Τα H8/H10 synthetic history datasets αποδεικνύουν συγκεκριμένα history ledgers και synthetic facts, όχι ότι κάθε patient/staff-bearing sheet, roster, hidden range, formula cache, named range και VBA form είναι synthetic και συνεπώς keyed.

Η δομή τεσσάρων H10 υποψηφίων που ελέγχθηκε στις 15/11 έδειξε τρία distinct workbook fingerprints:
- recurring IDs candidate,
- native write probe candidate,
- CC source / read-only probe pair, τα οποία είναι byte-identical.

Τα τρία distinct artifacts είναι frozen evidence/reference files. Δεν επιλέχθηκε κανένα ως writable integrated baseline.

---

## 5. Test evidence — τι είναι το τελευταίο πραγματικά green

### Full repository regression

Η τελευταία πλήρης regression πριν από την προσθήκη του νέου VBA project inventory tool ήταν:
- 1,813 passed,
- 1 skipped,
- 26 warnings,
- exit 0.

Αυτό είναι source regression, όχι native Excel acceptance.

### Afternoon focused regression

Η πιο πρόσφατη επιβεβαιωμένη focused suite για το afternoon branch ήταν:
- 49 passed in 9.22s.

Είναι source-only. Δεν υπήρξε native installer run.

### Permanent recurring append-plan CI

Η αναφορά καταγράφει CI run 43314844755 με:
- 1,652 full Python tests σε Windows,
- 1,652 full Python tests σε Linux.

Δεν ήταν native Excel event append.

### Νέο VBA project inventory tool

Προστέθηκαν:
- Python/tools/audit_vba_project_inventory.py
- Tests/test_audit_vba_project_inventory.py
- Documentation/H10_VBA_PROJECT_INVENTORY_READONLY_RUNBOOK_20261115.md

Ο κώδικας προστέθηκε και σκληρύνθηκε μετά την τελευταία full regression των 1,813 tests. Τα νέα focused tests και η full regression μετά από αυτές τις αλλαγές δηλώνονται NOT RUN. Επομένως το 1,813 δεν μεταφέρεται αυτόματα στον νέο κώδικα.

---

## 6. Κανόνες για τα αποδεικτικά Excel

- Κάθε GREEN native artifact παραμένει frozen και χρησιμοποιείται μόνο για read-only inspection.
- Κάθε FAILED artifact παραμένει frozen diagnostic evidence.
- Δεν γίνεται overwrite, repair-in-place ή δεύτερη εγκατάσταση στο ίδιο candidate.
- Το production baseline Rehab_Center_System_v27_1.xlsm παραμένει immutable.
- Τα H8/H10 attendance records είναι synthetic demo facts, όχι πραγματικές παρουσίες.
- Το native Excel proof εκτελείται σε Windows με πραγματικό Excel. GitHub CI και mocked COM tests δεν το αντικαθιστούν.
- Πριν από native write απαιτούνται source tests, provenance check, fresh unique destination, exclusive copy, native save/close/reopen και independent readback.
- Δεν ανεβαίνουν πραγματικά patient data, ονόματα, raw schedules ή private workbooks σε public GitHub ή public logs.

---

## 7. Πρωτογενείς αναφορές ανά workstream

### Κεντρικά handoffs / status
- ONBOARDING.md
- CURRENT_HANDOFF.md
- AUTHORITATIVE_RESUME_20261112.md
- PROJECT_STATUS.md
- IMPLEMENTATION_MAP.md
- DECISIONS.md
- PROJECT_HANDOFF_EVIDENCE_2026-11-03.md
- PROJECT_HANDOFF_NEXT_STEPS_2026-11-03.md

### H2–H7
- CLOSE_DAY_NATIVE_Z_PROOF_2026-11-01.md
- DAY_AMENDMENT_NATIVE_AB_PROOF_2026-11-01.md
- H4_THERAPIST_HISTORY_REPORTS_CHECKPOINT_2026-11-02.md
- H4_THERAPIST_HISTORY_REPORTS_GREEN_2026-11-02.md
- CHECKPOINT_H4_THERAPIST_HISTORY_2026-11-03.md
- CHECKPOINT_H5_PATIENT_HISTORY_2026-11-03.md
- CHECKPOINT_H6_CLOSED_H7_STARTED_2026-11-03.md
- CHECKPOINT_H7_BF_MANUAL_ACCEPTANCE_AND_PROTECTION_NEXT_2026-11-04.md

### H8–H10
- H8_STATISTICS_DASHBOARD_OPERATOR_FEEDBACK_2026-11-07.md
- H8_SYNTHETIC_HISTORY_DEMO_PLAN.md
- H9_THERAPIST_ATTENDANCE_HISTORY_CONTRACT.md
- H10_THERAPIST_CHARTS_USER_REQUIREMENTS_20261108.md
- H10_NEXT_SESSION_READ_FIRST_20261108.md
- NATIVE_CC_READONLY_PROBE_20261112.md
- NATIVE_CC_SYNTHETIC_SAVE_REOPEN_20261112.md
- NATIVE_SYNTHETIC_RECURRING_ID_MIGRATION_20261112.md

### Real source / registration / afternoon
- REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md
- FORMAL_ENTITY_REGISTRATION.md
- REGISTRATION_MENU_CONTRACT.md
- AFTERNOON_PHYSIOTHERAPY_MEDICAL_DEFERRAL_POLICY_20261115.md

### Protection / architecture
- BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md
- WORKSHEET_PROTECTION_WRITE_PATH_MAP_2026-11-04.md
- INTEGRATED_TEST_BASELINE_MATRIX_20261115.md

### Recovery chronology
Η σειρά RECOVERY_* διατηρεί τις διαγνωστικές αποτυχίες και τα fixes των native candidates. Οι τελικές accepted καταστάσεις προκύπτουν από τα νεότερα Z / AB / X proofs και τα αντίστοιχα CURRENT_HANDOFF appendices, όχι από την πρώτη εμφάνιση ενός παλιού «next step» σε ενδιάμεσο checkpoint.

---


## 7.1 Σημεία τεκμηρίωσης που χρειάζονται ενοποίηση

Η απογραφή εντόπισε ορισμένες ασυμφωνίες ονοματοδοσίας/κατάστασης ανάμεσα σε παλιότερες αναφορές. Δεν τις διορθώνουμε σιωπηρά:

1. **H2 Z filename:** το INTEGRATED_TEST_BASELINE_MATRIX_20261115.md αναφέρει UNIFIED_CLOSE_DAY_VERIFIED_PREVIEW_20261025_Z.xlsm, ενώ το CLOSE_DAY_NATIVE_Z_PROOF_2026-11-01.md και το CURRENT_HANDOFF.md αναφέρουν UNIFIED_CLOSE_DAY_HISTORY_20261101_Z.xlsm. Μπορεί να είναι διαφορετικά artifacts με ίδιο suffix. Χρειάζεται αντιστοίχιση με το τοπικό frozen-file ledger πριν επιλεγεί ένα ως canonical.
2. **H3 AB filename:** η acceptance matrix αναφέρει UNIFIED_CLOSE_DAY_RECOVERY_20261101_AB.xlsm, ενώ το DAY_AMENDMENT_NATIVE_AB_PROOF_2026-11-01.md αναφέρει UNIFIED_DAY_AMENDMENT_20261101_AB.xlsm. Δεν θεωρούνται ταυτόσημα χωρίς hash/path reconciliation.
3. **H9 test scope:** το H8 synthetic history plan αναφέρει 13 passed για H9 model/read-only loader, ενώ η H9 contract καταγράφει 7 passed για pure domain και στη συνέχεια προσθέτει store-reader tests ως pending. Χρειάζεται το ακριβές combined test log για να κλείσει η αμφιβολία.
4. **H7 BI vs BF:** BI είχε persisted install/protection audits green, αλλά η analytics UX απορρίφθηκε. BF είναι το μεταγενέστερο functional integration acceptance. Το BI δεν πρέπει να εμφανίζεται ως ισοδύναμο με BF.
5. **H10 visual vs functional acceptance:** τα CC screenshots αποδέχθηκαν δύο συγκεκριμένες οθόνες. Δεν αποδεικνύουν όλα τα period buttons, το πλήρες management XLSX export, κάθε report path ή real pilot readiness.

Αυτά είναι documentation reconciliation tasks, όχι απόδειξη ότι τα αντίστοιχα λειτουργικά tests απέτυχαν. Μέχρι να συμφωνηθούν τα ακριβή filenames / logs, διατηρούμε και τις δύο πρωτογενείς αναφορές και δεν επαναχρησιμοποιούμε κανένα artifact.

## 8. Προτεινόμενη σειρά από εδώ και πέρα

1. **Κλείσιμο της απογραφής VBA**: run των δύο focused auditor test files, full repository regression και read-only native inventory των τριών distinct VBA projects. Καμία αλλαγή σε frozen workbooks.
2. **VBA / installer ownership map**: αντιστοίχιση των components, UserForms, write surfaces και Python installers προς τα ήδη accepted workflows και την acceptance matrix.
3. **Ορισμός full synthetic fixture**: σχεδιασμός deterministic synthetic cohort και πλήρης cross-sheet identity/provenance audit πριν δημιουργηθεί οποιοδήποτε integrated XLSM.
4. **Κλείσιμο μηχανισμών με σαφή scope**: permanent event append/rollback/UI, H9 attendance writer/calendar, outpatient conflict/workload verification, afternoon save/medical deferral, global multi-absence optimizer.
5. **Release hardening**: fine-grained worksheet protection, 15-day patient progress workflow, final management XLSX privacy/visual acceptance.
6. **Real-source trial** μόνο αφού υπάρξει reviewed private source, explicit identity/slot approvals, supported native importer, independent readback και ξεχωριστή operator acceptance.

Δεν προτείνεται να ξαναγίνουν ήδη accepted tests χωρίς νέο code change ή συγκεκριμένη νέα ένδειξη. Κάθε νέο test πρέπει να κλείνει συγκεκριμένο gap, όχι να επαναλαμβάνει απλώς το παλιό proof.

---

## 9. Τελικό συμπέρασμα

Το Diomil δεν ξεκινά από το μηδέν. Έχει σημαντικό, πραγματικά αποδεδειγμένο κορμό καθημερινής λειτουργίας και ιστορικού. Οι αναφορές H2–H8 τεκμηριώνουν αρκετές ουσιαστικές native παραδόσεις.

Ταυτόχρονα, δεν είναι ακόμη ολοκληρωμένη και αποδεκτή ενιαία εφαρμογή. Οι πιο μεγάλες διαφορές ανάμεσα στο σημερινό αποτέλεσμα και στον τελικό στόχο είναι:
- full synthetic integrated workbook,
- genuine therapist attendance history,
- permanent recurring event transaction/rollback,
- real-source trial importer,
- full worksheet protection,
- global replacement optimization,
- afternoon operational integration,
- 15-day patient progress.

Αυτή είναι η σημερινή συγκεντρωτική εικόνα με βάση τις καταγεγραμμένες αποδείξεις της 15/11/2026. Οι επιμέρους αναφορές παραμένουν η πηγή αλήθειας για κάθε ακριβές test output, artifact και failure branch.


---

## 11. 2026-11-15 follow-up — native VBA inventory

The native read-only VBA inventory subsequently completed for three distinct H10 candidate projects. Two contain 67 components each; the CC candidate contains 66. The fourth candidate is byte-identical to the CC candidate. All distinct candidate hashes remained unchanged and no workbook was modified. Manual Calculation mode was unavailable in the local Excel COM environment and was recorded as a warning.

Thirty named application components have been mapped to tracked Python generators/installers. Thirty-seven components remain unresolved by source ownership: ten generic standard modules and 27 worksheet/document modules. The 66-component candidate is missing one zero-code-line document module whose exact identity is not recoverable from the captured Windows console encoding.

The detailed mapping and next operator gate are in Documentation/H10_VBA_OWNERSHIP_AND_WRITE_SURFACE_MAP_20261115.md. The next required proof is the sanitized, read-only VBA write-surface audit for the three distinct candidates. This inventory does not certify any candidate as the complete synthetic integration workbook and does not authorize changes to frozen workbooks.


## 12. 2026-11-15 write-surface audit follow-up

The first uploaded write-surface audit safely opened all three distinct candidate workbooks read-only, with macros/events disabled, and verified unchanged source hashes. It did not complete: the prior 200-match per-component output cap truncated modTherapistHistoryReport, which contains 265 matched source lines. The audit returned INCOMPLETE for all three candidates and is not accepted as the final write-surface proof.

The auditor was updated to 500 matches per component / 5,000 total, and its value-property classifier now distinguishes direct assignment from read/access. New regression coverage protects the known 265/1,396 match counts and read/write distinction. The operator must pull the latest branch, rerun focused tests and full regression, and regenerate the three read-only write-surface reports before source ownership reconciliation can be closed.


## 13. 2026-11-15 — completed native VBA write-surface evidence

The final uncapped write-surface report completed for all three distinct candidates in read-only mode, with macros and Excel events disabled. All three reported unchanged source hashes and no workbook modification. Component counts: 67, 67, and 66. Each produced 1,396 candidate source-pattern matches without truncation.

There are 35 components with at least one matched source-pattern line per candidate. Reconciliation with the native inventory requires normalizing two worksheet-module name encoding aliases. The remaining 32 components without a matched operation pattern consist of 23 zero-code components and 9 non-empty components requiring targeted source review. Absence of a pattern match is not proof of no writes.

The 1,396 matches are not confirmed runtime writes. Per-workbook category counts: worksheet references 80; value/formula accesses 373; direct value/formula writes 282; range/structure mutation candidates 111; format/other candidates 460; visibility/action binding 30; workbook saves 26; unprotect 17; protect 17.

Full regression completed: 1,824 passed, 1 skipped, 26 warnings. The earlier "37 unresolved components" note is superseded; use 32 after accounting for the two worksheet-module name encoding aliases.


### Reconciliation count clarification

The 32-component no-match reconciliation (23 zero-code + 9 non-empty) is grounded in the 67-component inventory candidate and applies to the two 67-component reports. The CC read-only probe has 66 components and the same 35 matched components, so its arithmetic remainder is 31, not 32. Its candidate-specific component inventory has not yet been reconciled item-by-item; do not project the 23/9 breakdown onto that 66-component workbook. The two worksheet-module display-name aliases are normalized only when matching against the 67-component inventory.


## Targeted VBA source review — 2026-11-15

Tracked source review clarifies five named components previously absent from the worksheet-write detector. The registration menu, outpatient entry macro, outpatient schedule manager and scheduling-conflict popup are navigation/read-only UI or explicit confirmation surfaces; the outpatient form delegates actual writes to the Python preview/authoritative-commit worker. The double-booking review module does perform external file I/O and process orchestration, including temporary request/response files and authoritative workbook close/reopen coordination. It was not a no-op; it was outside the previous worksheet-only detector's scope.

Source-contract regression tests exist for all five components. They do not replace native end-to-end acceptance.

Generic Module5, Module6 and two non-empty worksheet/document modules remain unresolved. The auditor is extended to output procedure names and worksheet names for no-pattern components, without VBA source text, and to detect external file I/O, process launch and workbook close. A fresh operator run is required before this subtask can close.


## Latest continuation gate — 2026-11-15

The VBA auditor was further extended after the first targeted source review:
- emits NO_PATTERN_MATCH_COMPONENT records with escaped component name, type, line count, worksheet name when applicable, and procedure identifiers only; it never prints source text;
- detects external file create/write/delete/read, process launches, and ThisWorkbook.Close in addition to worksheet and workbook write surfaces;
- excludes procedure declaration lines from operation matches;
- recognizes .Run(...) calls with or without whitespace.

New focused regression tests cover procedure inventory, external I/O/process/lifecycle classification, and declaration-line exclusion. These latest changes have not yet been executed on the operator's Windows/Excel machine. The previous 1,824-pass full regression and completed native audit predate these latest changes and must not be presented as validation of them.

Next operator gate: pull branch work/afternoon-medical-scope-20261115; run Tests/test_audit_vba_write_surfaces.py and the existing integrated-structure/VBA-inventory tests; run full pytest; then rerun the sanitized audit on the same three distinct H10 candidates. Upload the focused/full regression logs and final write-surface report. The new report should expose procedure identifiers for the remaining unmatched non-empty components and separately classify the external file/process operations in modDoubleBookingReview. Preserve ReadOnly=True, disabled macros/events, unchanged source hashes and no workbook modifications.


## Operator validation received — 2026-11-15 19:34 UTC

Evidence bundle: DIOMIL_VBA_SOURCE_REVIEW_20261115_193409.txt

Results:
- Focused auditor/inventory/structure tests: 24 passed.
- Full pytest: 1,829 passed, 1 skipped (symlinks unavailable), 26 warnings, 29.64 seconds.
- Native read-only VBA write-surface audit completed for all three H10 candidates: UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm, UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm, and UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm.
- Each audit reports 1,569 matched operations, SOURCE HASH UNCHANGED: YES, and NO WORKBOOK WAS MODIFIED.
- Read-only safeguards reported: ReadOnly=True, macros disabled before open, Excel events disabled.

The same eight no-pattern components appear in each candidate: ThisWorkbook document module (19 lines; Workbook_Open/Activate/Deactivate/SheetActivate/BeforeClose), worksheet module for PATIENT_PLANNER (25 lines; Worksheet_Change), Module6 (26 lines; NormalizeTimeMS and NormalizePresenceMS), Module5 (1 line; no procedure), modRegistrationMenu, modOutpatientSchedule, modOutpatientScheduleManager, and modSchedulingConflict. The four named UI/navigation components were source-reviewed separately and are not direct worksheet-write procedures. The workbook event module, PATIENT_PLANNER Worksheet_Change, and Module6 still need explicit ownership/behavior review; Module5 should be checked for whitespace/attribute-only/empty content before classifying as inert.

The modDoubleBookingReview external file/process/workbook-close behavior is now visible in the audit, confirming why the former worksheet-only scan was incomplete. Do not interpret the 1,569 pattern matches as 1,569 confirmed writes: the report includes accesses, reads, formatting/navigation, and writes. The classification is a static surface inventory, not runtime proof.

Next: review the three remaining behavior-bearing no-pattern components (ThisWorkbook events, PATIENT_PLANNER Worksheet_Change, Module6 time/presence normalization); classify Module5; then consolidate component ownership, side-effect type, target sheets/files/processes, and test/native evidence into the master write-surface map. No further workbook mutation or native write trial is authorized by this audit.


## Native VBA audit and regression results — 2026-11-15 19:34 UTC

Operator uploaded DIOMIL_VBA_SOURCE_REVIEW_20261115_193409.txt. Evidence reviewed:
- Focused tests: 24 passed in 2.50s.
- Full pytest: 1,829 passed, 1 skipped, 26 warnings in 29.64s (1,830 collected).
- Three distinct VBA candidates completed the read-only write-surface audit. Each reports ReadOnly=True, macros disabled before open, Excel events disabled, SOURCE HASH UNCHANGED: YES, and NO WORKBOOK WAS MODIFIED.
- Candidate hashes: RECURRING_IDS = 4df997b6a5e5b0b56b37127ee00dd8710cc26c7c118f1742d58c9969bcbe7afa; NATIVE_WRITE = fa9765da7079137bcb1ff935c3d08bc84b409c8335f9cd719f3583b89e4fc67d; CC_NATIVE_READONLY_PROBE = 1bd5e58e905f91acb55e5a706cd8b4563b5a98c170a9ec60fe76c95617420b13. Candidate counts are 67, 67 and 66 VBA components respectively.
- The separate native inventory reports Excel calculation mode could not be set to Manual because of a COM error. This is a warning, not a failed audit; retain it as an explicit limitation. ReadOnly and macro/event safeguards were confirmed.
- The write-surface audit detects 1,569 matched operations in each candidate. Treat this as a static pattern inventory, not 1,569 proven mutations or a runtime execution trace.

Unmatched-component investigation now has useful identities:
- The 19-line DOCUMENT_MODULE is the workbook event module: Workbook_Open, Workbook_Activate, Workbook_Deactivate, Workbook_SheetActivate, Workbook_BeforeClose.
- The 25-line worksheet document module maps to PATIENT_PLANNER and contains Worksheet_Change.
- Module6 (26 lines) contains NormalizeTimeMS and NormalizePresenceMS. These are helper procedures with no current write-surface pattern match; inspect their actual implementation/consumers before assigning a stronger classification.
- Module5 contains one code line and no recognized procedure.
- modRegistrationMenu, modOutpatientSchedule, modOutpatientScheduleManager and modSchedulingConflict are now identified by entry procedure names and their source-level roles are documented in the targeted review.
- modDoubleBookingReview is positively identified as an external side-effect orchestrator: temporary file create/delete/read/write, Python process launches and workbook close/reopen coordination. The scanner reports these explicitly (16 matches in the module), so it must be treated as a write/lifecycle owner even though it has no direct worksheet assignment in the entry module.

Important audit limitation: procedure names and pattern matches identify ownership candidates, not full semantics. The two event procedures, Module6 helpers and one-line Module5 still require implementation-level review and caller/callee tracing. Do not mark the VBA ownership map fully closed yet. The inventory emits component names in escaped Unicode where applicable; no source text or cell values were printed.

Next step: complete the remaining source-level trace for Workbook events, PATIENT_PLANNER.Worksheet_Change, Module6 and Module5 using tracked VBA generator/source where available, and record unresolved native-only code honestly. Do not modify source workbooks. The regression and read-only audit gate is green; proceed with documentation and static ownership analysis, not production workbook changes.


## 2026-11-15 — H10 native VBA audit update

The latest operator evidence bundle is DIOMIL_VBA_SOURCE_REVIEW_20261115_193409.txt:
- Focused tests: 24 passed.
- Full regression: 1,829 passed, 1 skipped, 26 warnings.
- Three distinct H10 candidate workbooks completed the expanded native read-only VBA write-surface audit.
- Each candidate produced 1,569 static source-pattern matches. ReadOnly=True, macros disabled before open, Excel events disabled, SHA-256 unchanged, and no workbook modified.
- Calculation mode could not be set to Manual due to a COM error; recorded as a non-blocking limitation.

The audit exposes external file I/O, process-launch and workbook-close lifecycle behavior in modDoubleBookingReview. It also identifies eight components with no operation-pattern match per candidate: ThisWorkbook events, PATIENT_PLANNER.Worksheet_Change, Module6 (NormalizeTimeMS/NormalizePresenceMS), Module5 (one line/no procedure), and four named UI/navigation entry modules. The four named UI modules were matched to tracked source and reviewed. The four native-only components still require behavior-level review.

Important: 1,569 matches are static source-pattern locations, not confirmed runtime writes. The VBA inventory/audit gate is complete; source ownership mapping and semantic review remain partial. The full synthetic integrated workbook and actual patient-source admission are still separate, unaccepted milestones.


## 2026-11-15 — source-redacted procedure profiling added

The read-only VBA auditor now emits additional metadata for components with no operation-pattern match:
- per-component SHA-256 fingerprint;
- per-procedure SHA-256 fingerprint;
- procedure line range and body-line count;
- statement-category counts only (assignments, branches, loops, date/time operations, event/Target references, calls, other).

It does not emit VBA source lines, literals, patient names, cell values, or formulas. This is intended to compare the embedded implementations across the three native candidates and provide a safe structural profile for the four remaining native-only components (ThisWorkbook events, PATIENT_PLANNER.Worksheet_Change, Module6, Module5).

New regression tests cover source redaction, stable procedure hashes, profile counts and the emitted metadata contract. These latest profile changes have NOT yet been executed locally. The previously reported 24 focused passes / 1,829 full passes validate the preceding auditor revision, not this new profile revision. A fresh focused/full/native run is required before relying on the new profile output.


## 2026-11-15 19:50 — H10 procedure-profile acceptance and event source reconciliation

Operator evidence `DIOMIL_VBA_PROFILE_BUNDLE_20261115_195014.txt`: 27 focused tests passed; 1,832 full tests passed, 1 skipped, 26 warnings. Three distinct H10 native read-only XLSM audits passed, with 1,569 static operation matches each, matching nonempty component inventories, unchanged SHA-256s and no workbook modifications. Detailed acceptance is in `Documentation/H10_VBA_NATIVE_PROCEDURE_PROFILE_ACCEPTANCE_20261115.md`.

The tracked `Python/rehab_excel/application_shell_vba.py` workbook-event installer was compared with the native procedure SHA-256s: all five `ThisWorkbook` event bodies matched exactly, so ownership for that event-call scope is now confirmed. The regression assertion `Tests/test_h10_native_workbook_event_hash_contract.py` was committed but is not covered by the preceding green test count. `Module5` is empty in the native code-module output despite CountOfLines=1. The remaining native-only event/normalizer semantics of `PATIENT_PLANNER.Worksheet_Change` and `Module6` must not be inferred. No frozen candidate was modified; the integrated synthetic workbook and real-source admission remain unaccepted.


## 2026-11-15 — static VBA write-surface cross-candidate verifier committed

New code: `Python/tools/verify_h10_vba_profile_reports.py`, `Tests/test_verify_h10_vba_profile_reports.py`. The report verifier parses sanitized UTF-16/UTF-8 VBA audit outputs and checks three unique named H10 candidates, safety and hash markers, expected component counts (67/67/66), complete operation totals, and cross-candidate component/fingerprint/operation/profile equality. Its new regression suite uses synthetic report fixtures only.

Independently parsed uploaded native report `DIOMIL_VBA_PROFILES_20261115_195014.txt`: each of three candidates has 44 nonempty component records and 1,569 operation-pattern matches. Category counts per candidate: 282 direct Value/Formula assignment candidates, 111 range/structure candidates, 17 Protect, 17 Unprotect, 26 workbook saves, 4 workbook closes, 28 external launches, 91 external file write/delete, 46 external file read, plus 373 value accesses, 80 worksheet references, 464 format/other and 30 visibility/action candidates. All three reports agree in these category counts. This is static-source-pattern evidence, not operational runtime mutation proof.

Full detailed matrix: `Documentation/H10_VBA_NATIVE_PROCEDURE_PROFILE_ACCEPTANCE_20261115.md`. No source workbook or clinical data was modified or committed. The new verifier/tests were not part of the preceding 1,832-pass full regression and must be validated before claiming their acceptance. Do not require repeat native workbook runs solely for documentation; run the new verifier against the already uploaded sanitized profile log on the operator's next ordinary test gate if necessary. Remaining native event-semantic work: `PATIENT_PLANNER.Worksheet_Change` and `Module6` normalization rules/callers.


## 2026-11-15 — local verification of existing H10 evidence, no new native Excel needed

The exact uploaded `DIOMIL_VBA_PROFILES_20261115_195014.txt` was parsed offline with the GitHub verifier's comparison logic: **GREEN** for 3 distinct workbooks, 44 nonempty VBA components each, 1,569 categorized matches each and 12 procedure-profile hashes each, with equality of recorded components/operations/profiles across all three. This validates the archived profile report, not any live write behavior. The earlier 1,832-pass regression predates the newer source changes.

Added read-only auditor `POSSIBLE_CALL_EDGE` metadata collection for no-pattern components (including `PATIENT_PLANNER.Worksheet_Change` and `Module6`), with quoted strings/comments excluded and only declared procedure identifiers emitted. Added synthetic tests in `Tests/test_audit_vba_write_surfaces.py`. These newest code/tests have not yet been operator-tested.

Per-workbook write-surface priorities, from the archived report: `modTherapistHistoryReport` (74 direct assignment candidates), `modPatientRegistrationBridge` (42), `modDailyChangeLedger` (29), `modGlobalReplacementAction` (25), `modReplacementOperator` (20), legacy `Module4` (18). `modDoubleBookingReview` is an external file/process/workbook-lifecycle owner despite having 0 direct worksheet assignments. Detailed caveats and sheet references are in `Documentation/H10_VBA_NATIVE_PROCEDURE_PROFILE_ACCEPTANCE_20261115.md`.

Immediate next stage: no need to repeat old inventory. When the next operator/native gate is genuinely required, run the new focused tests and read-only auditor on the three frozen workbooks, capture `POSSIBLE_CALL_EDGE` records, verify unchanged hashes, and trace any identified event/normalizer callees. Do not claim the new tests green before their run. Then continue to full-cohort synthetic integration design, keeping actual patient intake separate.


## 2026-11-15 — call-edge reconciliation now covered by tests

The targeted, source-redacted `POSSIBLE_CALL_EDGE` output was integrated into `Python/tools/verify_h10_vba_profile_reports.py`. The verifier now compares candidate references between all three distinct workbooks in addition to component metadata, operation matches and procedure fingerprints. New synthetic tests in `Tests/test_verify_h10_vba_profile_reports.py` exercise matching call edges and an intentionally divergent third-candidate edge. Quoted string literals and VBA comments are excluded during the native candidate extraction; the output is still only a *possible reference*, not proven runtime call execution.

**Latest change test status:** the new call-edge extraction, verifier extension and regression assertions have been committed but have not yet been executed on the operator's Windows source checkout. The previous 1,832-pass full suite is not proof for these new commits. The previously uploaded native profile report did not yet include call-edge records, so it cannot establish the unknown PATIENT_PLANNER event's actual call targets.

No additional Excel workbook was opened or saved during this GitHub-only update; previously frozen native evidence remains unchanged.


## 2026-11-15 — verifier hardening and one-command Windows call-edge gate

Changes implemented on `work/afternoon-medical-scope-20261115`:

1. Fixed a real parser defect: the `POSSIBLE_CALL_EDGE` regex accidentally escaped `\\s`, `\\S` and `\\d` twice, preventing correct recognition. The corrected expression has an explicit parsing regression test.
2. The native read-only auditor now searches **all components with existing operation matches** for potential references to `NormalizeTimeMS` and `NormalizePresenceMS`, not just the eight no-pattern components. Existing no-pattern call-edge records are retained. These are **candidate identifier references**, not proof of runtime invocation.
3. Corrected normalized no-pattern `component_sha256`: its line separator is now an actual LF, not a literal backslash plus `n`. **Previous native source-redacted component fingerprints are historical fingerprints from the old algorithm**; they are comparable within that original report, but must not be compared directly against newly generated component fingerprints. Per-procedure hashes were already normalized with LF and are not affected by this component-hash correction.
4. The verifier now refuses malformed operation/profile/call-edge records and compares each component's declared `matched_operations` count to the number of parsed operation rows, in addition to global counts, safety attestations and three-candidate equality. Regression cases for malformed input and mismatch were added.
5. Created `Scripts/run_h10_vba_call_edge_gate.ps1`. It enforces the correct branch, existence of three frozen reference XLSMs, fresh unique output logs, focused pytest, full pytest, read-only native auditor, 3/3 successful audit/hash/no-modification markers and the offline three-candidate verifier. It bundles results into a single upload file. The script never saves or edits a source XLSM and does not automatically retry failed candidates.

Archived uploaded report `DIOMIL_VBA_PROFILES_20261115_195014.txt` was inspected offline again: exactly three completed audit blocks, each with 1,569 parsed operation rows, 12 procedure profiles and eight no-pattern components. This is a read-only check of **old evidence**, not proof for the newer scanner/verifier or the new one-command gate.

**Validation status:** All newest source changes and PowerShell runbook are committed, but no post-change Windows PowerShell/pytest/native run has yet been observed. Do not describe the new version as accepted until the next operator report. The previously accepted 1,832-pass regression and three preserved source hashes remain valid for their earlier revision. The next required native evidence is not an integration installation; it is the targeted call-edge proof needed to resolve `PATIENT_PLANNER.Worksheet_Change` and `Module6` helper callers.

### Exact operator command when the next gate is ready

With Excel closed, from the repo directory:

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull failed" }
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\run_h10_vba_call_edge_gate.ps1
```

Upload the printed `DIOMIL_H10_CALL_EDGE_BUNDLE_*.txt`. The PowerShell script stops on failed source regression, absent candidates, incomplete native results, changed workbook fingerprints, or cross-candidate differences. It does not grant permission for a writable synthetic demo or for any real patient migration.

After successful evidence review, continue the new fully synthetic integrated fixture design, maintaining explicit separation from frozen native candidates and real admission paths.


## 2026-11-15 20:31 — H10 call-edge gate accepted; event/validator ownership narrowed

Primary operator artifact: `DIOMIL_H10_CALL_EDGE_BUNDLE_20261115_203111.txt`.

**Observed Windows results:**
- Focused tests: **38 passed in 1.94s**.
- Full repository regression: **1,854 passed, 1 skipped, 26 warnings in 31.08s**; no failures.
- The one skip remains `Tests/test_afternoon_sheet_preview.py:68`, because symlinks are unavailable in the Windows environment.
- Native read-only VBA write-surface audit completed for all three distinct H10 candidates. Each retained `ReadOnly=True`, macros/events disabled, `SOURCE HASH UNCHANGED: YES`, and `NO WORKBOOK WAS MODIFIED`.
- Each candidate still reports 1,569 static operation-pattern matches. Cross-candidate verifier: **GREEN**, three distinct workbooks, 44 nonempty component metadata records, 12 procedure-profile hashes and 9 candidate call-edge references. The pattern count is not a runtime write count.

**Call-edge detail (same in all three candidates):**
1. Five `ThisWorkbook` event calls: Open/Activate → `EnterApplicationShell`; Deactivate/BeforeClose → `ExitApplicationShell`; SheetActivate → `KeepApplicationShell`. These five calls were already matched at exact body-hash level to tracked `application_shell_vba.py::_EVENT_SPECS`.
2. `PATIENT_PLANNER.Worksheet_Change` has three branches, each referencing `CheckSingleAssignment` (native lines 16, 19, 22).
3. `modSchedulingConflict.ConfirmTherapistDoubleBooking` references `ConfigureConflict`.

The native component `Module11` declares `CheckSingleAssignment`, plus `BuildBusyMapsForCheck`, `GetPresentTherapists`, `GetTherapistLoad` and `GetSlotList`. Its 30 detected operation rows are worksheet references and cell-value/formula **accesses** only; no direct assignment or range/structure mutation candidate was detected in this component. Its literal worksheet references are `PATIENT_PLANNER`, `THERAPIST_ATTENDANCE` and `SETTINGS`. This supports classifying it as a read-side single-assignment/availability validation subsystem at the static-pattern level; it is not a complete control-flow proof.

`Module6` exposes `NormalizeTimeMS` and `NormalizePresenceMS`, with identical procedure fingerprints in the three candidates. No candidate VBA call-edge to either helper appeared in this report. This does **not** rule out formula/UDF references, external COM calls, or other invocation mechanisms not represented as a VBA identifier call.

**Next read-only refinement committed after this accepted bundle:**
- The auditor now searches for candidate references to *all* declared VBA procedures in every component with matched operation rows, rather than limiting those components to the two Module6 normalizers. This should reveal Module11's own helper call graph and other cross-module dependencies.
- The structural OOXML audit now counts worksheet-formula references to `NormalizeTimeMS` and `NormalizePresenceMS` without printing formula text, values, identities or cached values. Excel string literals are stripped before matching to avoid false positives.
- The one-command Windows gate now includes this formula scan on the same three frozen candidates and the focused regression for its new privacy-safe parser.

**Current status:** the just-reported 38/1,854 GREEN results validate the preceding revision, not the expanded call-graph/UDF-formula changes committed afterwards. Those newest changes require a fresh focused/full/native read-only gate. No frozen workbook has been modified or selected as a writable integration base.

Next product phase remains the whole-workbook demo. The real-patient source/review workbook and separate eligible pilot XLSM remain a distinct, operator-held admission gate. Do not merge real patient data into a synthetic fixture or treat the H10 candidates as writable targets.


## 2026-11-15 20:40 — expanded all-procedure call graph and formula-only UDF scan prepared

Following the accepted 20:31 bundle, two further **read-only** analysis improvements were committed:

1. `Python/tools/audit_vba_write_surfaces.py` now emits source-redacted `POSSIBLE_CALL_EDGE` candidates against **all declared VBA procedures** from every component with matched operation rows, not only the two Module6 normalizers. This is intended to expose Module11 helper calls and other cross-module edges. Edges are identifier references after quoted strings/comments are removed; they are not runtime execution proof. The synthetic tests now assert cross-procedure discovery and that the previous normalizer-only filter is absent.
2. `Python/tools/audit_integrated_workbook_structure.py` now counts formula references to `NormalizeTimeMS` and `NormalizePresenceMS` in the formula view, per sheet and per workbook, without printing formulas, values, cached results or identities. Quoted Excel string literals are removed before matching. Tests cover actual function calls and false-positive text constants.

`Scripts/run_h10_vba_call_edge_gate.ps1` was updated to include the new focused test file and run the read-only structure/UDF formula scan over the same three frozen H10 candidates. It requires 3/3 completed audits, unchanged source hashes and no-modification attestations, then adds the structure/UDF log to the single output bundle.

**Validation boundary:** these new commits are after the accepted 38-focused / 1,854-full result and have not yet been run by the operator. They must not inherit that earlier GREEN. The next local action is one execution of the updated PowerShell gate with Excel closed. The new script still does not save or modify any XLSM.

After this proof, proceed to the full synthetic integration workbook design/build. The real-source file named `DIOMIL_ΕΛΕΓΜΕΝΗ_ΜΕΤΑΦΟΡΑ_ΠΗΓΗΣ_ΠΡΟΣ_MASTER_20261112.xlsx` was not found among the currently available Project attachments; do not guess its local location or commit it. When the real-data admission milestone is reached, ask the operator to supply the reviewed source privately and follow `REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md`. The A:T column mapping already supplied in the conversation is retained as the operator's source-schema reference; the reclining-treatment replacement rule still requires the operator's fresh clinical instructions at the point of implementation.


## 2026-11-15 20:50 — expanded H10 call-edge gate accepted; private source inputs identified

**Operator evidence bundle:** `DIOMIL_H10_CALL_EDGE_BUNDLE_20261115_205003.txt` (private conversation upload; do not copy the bundle or workbook data into the public repository).

### Accepted validation
- Focused tests: **45 passed in 2.47s**.
- Full repository regression: **1,856 passed, 1 skipped, 26 warnings in 30.07s**; exit was successful.
- The sole skip is `Tests/test_afternoon_sheet_preview.py:68` because symlinks are unavailable in the Windows environment.
- Expanded cross-candidate VBA evidence verifier: **GREEN**, 3 distinct workbooks, 44 nonempty component metadata records, 12 procedure-profile hashes and 1,027 possible static call-edge references.
- Each of the three distinct H10 candidates completed the read-only VBA write-surface audit with `ReadOnly=True`, macros/events disabled, `SOURCE HASH UNCHANGED: YES`, and `NO WORKBOOK WAS MODIFIED`.
- Each candidate reports 1,569 static operation-pattern matches. These are lexical candidates, not runtime write counts.
- Read-only workbook structure scan completed for all three candidates; source hashes unchanged and no workbook was modified.
- Worksheet-formula references to `NormalizeTimeMS` and `NormalizePresenceMS`: zero in all three inspected candidates. This rules out those worksheet-formula references only, not every possible invocation mechanism.
- The 1,027 call edges are static identifier references, not proof that every branch executes at runtime.

### Product target and private intake boundary
The operator's target is one functional integrated Excel demo using the reviewed patient/treatment source and the previously supplied A:T field mapping. Two private conversation uploads are available for local inspection: the 10-11-26 patient treatment programme workbook and the 12-11-26 identity-review/proposals workbook. They are not repository assets and must never be committed, attached to public issues, or copied into public logs.

The source programme sheet `ΠΡΟΓΡΑΜΜΑ` has the operator's A:T fields; U/V carry additional afternoon and pool fields. Preserve these as separately mapped fields rather than silently dropping them. The identity workbook is a review/proposal artifact, not approval: every patient identity remains pending, exact-name suggestions are not confirmations, and flagged source overlaps/notes remain on hold. Do not import or activate real patient rows until identity decisions and conflict/source holds have been explicitly resolved. No live import or real-data XLSM has been produced by this inspection.

The previously referenced filename `DIOMIL_ΕΛΕΓΜΕΝΗ_ΜΕΤΑΦΟΡΑ_ΠΗΓΗΣ_ΠΡΟΣ_MASTER_20261112.xlsx` is not the name of either available upload. Do not ask for a re-upload of the two available files. If a later stage requires that distinct reviewed/approved artifact, request it only then.

### Next work
1. Keep all private source files local and read-only.
2. Reconcile the private source schema and identity-review contract against the tracked trial admission/preflight/projection code.
3. Build a synthetic, source-shaped fixture and test the complete A:V parser/projection without real patient data.
4. Close any remaining native-only VBA ownership questions and define the fresh integration candidate from a trusted source; never write into any frozen H10 proof workbook.
5. Only after synthetic source-to-workbook roundtrip is proven, run the separate private identity/conflict admission gate and request operator decisions for unresolved clinical/source holds.


## 2026-11-15 — A:V real-source demo schema design

Added `Documentation/TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md` as the current source-to-demo design contract. It records the operator's A:T mapping, the additional U/V afternoon/pool fields, the boundaries of the existing read-only admission/import-plan/weekly-projection modules, and the required architecture for source alignment, patient registry projection, full recurring schedule projection, synthetic end-to-end tests and later native Excel acceptance.

Important unresolved rules are explicit rather than guessed: blank weekday semantics; missing A context handling; U/V operational scope; and fresh operator instructions for reclining replacement behavior at implementation time. The private identity review remains unapproved; no real-patient import or native integrated demo is authorized. Keep source workbooks and all patient-level facts outside GitHub.

 
## 2026-11-15 21:15 — source-alignment audit and cross-platform CI accepted

Added:
- `Python/rehab_excel/trial_source_alignment.py`: read-only comparison of original A:V programme rows, identity review rows and weekly source-cell references. It emits aggregate counts/safe blocker codes only and verifies both input hashes remain unchanged.
- `Python/tools/audit_trial_source_alignment.py`: local private-input CLI; never prints patient names, IDs, room/doctor values, raw source cells or local paths.
- `Tests/test_trial_source_alignment.py`: synthetic tests for exact provenance, identity/context mismatches, wrong source rows/cells, unsupported specialties, missing source times, reclining's time-only source reference, unchanged files and CLI privacy.
- `Documentation/TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md`: A:V field contract and complete source-to-demo architecture.

**GitHub Actions run 43751221120 — GREEN**, commit `94618fa6402c3cfbf028dbcc32b5f694377df754`:
- Focused synthetic safety: **200 passed** on Ubuntu and Windows.
- Full regression: **1,869 passed** on Ubuntu (15 warnings) and Windows (14 warnings).
- No failed or skipped tests in the reported test summaries.

A separate local/private cross-check of the currently uploaded source and review pair found matching patient source-row/name/context provenance and matching weekly source-cell references for the staged review rows. Some A-column ward/doctor contexts are blank and remain blank; no forward-fill was inferred. This check does not approve identities or prove the semantic correctness of proposed weekdays/times/providers.

**Still blocked:** every PatientID confirmation remains pending; suggestions are not approvals; source notes and overlap holds remain unresolved. No real-patient import, native writer or integrated real-data XLSM has been run or accepted. The new alignment CLI is read-only.

**Next engineering slice:** bind the original A:V source hash into the private import plan; build a complete patient-registry plus weekly-treatment projection and synthetic roundtrip; retain U/V as explicitly unmapped until their operational contract is decided. Before interpreting blank weekday cells, obtain the operator's rule. Do not enable import from the existing proposals.

 
## 2026-11-15 — operator decision: weekday-only schedule and blank day cells

Operator clarification: **there are no Saturday/Sunday treatments; the operating treatment week is Monday through Friday (ΔΕ–ΠΑ).** For source import interpretation, when a treatment's corresponding time/value is present but its weekday cell is blank, interpret that treatment as recurring Monday–Friday. Never generate Saturday or Sunday sessions. A missing treatment time/value means no session is created; do not create a weekday slot from a blank day cell alone.

Implementation must keep this as an explicit source-import rule (pass `implicit_daily=True` only for the relevant source mapping when a treatment value exists and its day cell is blank), rather than changing the general parser default for unrelated callers. Add synthetic regression coverage for blank-day + present-time => exactly ΔΕ, ΤΡ, ΤΕ, ΠΕ, ΠΑ and zero weekend entries; blank-time + blank-day => zero entries. Preserve explicitly supplied weekday patterns unchanged.

## 2026-11-15 — A:V weekday candidate projection (source-only)

The pure module `Python/rehab_excel/trial_source_weekday_projection.py` now maps individual A:V-shaped rows into candidate `SourceWeeklyEntry` records, reusing `rehab_core.programme_mapping.make_weekly_entries`. It has no workbook input/output and no identity auto-approval. Only a separate, explicitly confirmed PatientID may be supplied by the future private admission flow.

- A treatment clock with a blank weekday field becomes exactly five recurring weekdays (ΔΕ–ΠΑ), using `implicit_daily=True` locally without changing the general parser default.
- Blank clock creates no slot; nonblank days/provider alongside a blank clock become a privacy-safe missing-clock blocker.
- Explicit weekday/time groups keep their own exact weekday and clock; invalid patterns and weekends are blocked.
- Reclining F and afternoon/pool U/V are retained as blockers until their dedicated operational rules exist. If any field blocks, the complete row returns no admissible candidates, not a silently truncated programme.
- Source cell coordinates and clinical review provenance remain necessary. The module does not establish identity confirmation, overlap resolution, native importer functionality, or application acceptance.

Synthetic-only regression: `Tests/test_trial_source_weekday_projection.py`. These new tests are **not yet verified green** until their own CI result is checked. The previously accepted 1,869-pass CI applies to its earlier commit, not this addition.

**Immediate next integration work:** connect the source-row reader and explicitly confirmed private identity mapping to this projection; add the complete patient registry projection and bind the original programme SHA256 to the plan before any native Excel writer. The first requested deliverable remains a single operational demo for admissions, discharge, absences, replacements, therapists and students; H9/H10 analytics are not the immediate gate.


## 2026-11-15 — DEMO priority: source-bound patient registry + weekly programme (engineering checkpoint)

**Operator priority:** build ONE openable integrated practice DEMO for admission/discharge, patient/therapist absences and replacements, new therapists/students, restoring changes and day operations. Statistics, H9 attendance charts, 15-day patient progress and management analytics are subsequent/parallel, **not blockers for the first practice DEMO**. Do not start over with H2–H10 native proofs or repeat previously accepted gates without a code change or new evidence.

**Code delivered on `work/afternoon-medical-scope-20261115` after the already accepted 1,869-pass source-alignment CI:**

- `Python/rehab_excel/trial_source_weekday_projection.py`: pure, source-cell-preserving A:V row expansion. Populated clock + blank days maps to exactly ΔΕ–ΠΑ; blank clock creates no slot; explicit split weekdays/times retained; no weekends. The import-specific `implicit_daily=True` is never made a global parser default. Unsafe/incomplete inputs hold the entire source row, not a partial import. Reclining F and afternoon/pool U/V are **unmapped blockers**, not silently discarded. This is a proposed mapping only.
- `trial_import_plan.py`: optional `programme=` requires original source alignment and binds its SHA256 alongside review and eligible intake. The old two-input plan remains readable but **is not source-bound**; it may not be used by the new registry projection. A source-bound plan requires the programme again for `verify_sources` and blocks any changed/missing source.
- `Python/rehab_excel/trial_patient_registry_projection.py`: read-only `PATIENTS` A:E + optional type/MRN/doctor and separate `SourceContext` in-memory rows, with exact original A/B source row and review confirmation checks. Clinical `PatientType`, infectious yes/no, and any room/doctor/status must be independently provided by an authorized reviewer, never derived from names, section headings, ward/doctor context, or guessed formatting. Blank source context stays blank. Nothing writes patient data to GitHub.
- `Python/rehab_excel/trial_demo_source_projection.py`: combines that registry with existing `PROGRAM_WEEKLY_SLOTS` projection, requiring all weekly PatientIDs to exist in the explicitly classified inpatient registry. Outpatient weekly entries remain blocked under this inpatient-only schema rather than silently omitted.
- Synthetic tests: `Tests/test_trial_source_weekday_projection.py`, `Tests/test_trial_patient_registry_projection.py`, plus original programme SHA/legacy fail-closed cases in `Tests/test_trial_import_plan.py`. Tests use invented identities and do not open real source workbooks.

**Boundary:** This is a read-only, source-level implementation milestone; it is **not yet a native XLSM importer, integrated practice workbook, confirmed real patient migration, certified operational demo or clinical-data approval**. Do not use frozen H10 source/proof workbooks as writable destinations. CI result for each new commit must be confirmed separately; the older 1,869/200 GREEN cannot be inherited by new changes.

**Next development sequence:** independent source-bound projection fixture and exact weekly/registry reconciliation -> eligible disposable native workbook writer and independent Windows Excel readback -> integrate the already existing admission, discharge, replacement, therapist/student registration and in-session UX within the same file -> operator practical scenarios. Actual reviewed source and identity approval stay private; no approval means no real patient import. Keep demo repeatable via a clean, frozen starting copy. Historical resume/test instructions above this appendix remain chronology only where contradicted by this newer checkpoint.


## 2026-11-15 — DEMO A:V source versus admitted weekly parity GREEN

**Priority remains one integrated practice DEMO first**, covering patient admission/discharge, absences, replacements, new therapists/students and daily operations. History, H9/H10 statistics, 15-day patient progress and management reporting are later/parallel; they are not prerequisites for the practice workflow.

**Delivered:** `Python/rehab_excel/trial_demo_source_reconciliation.py` independently reconstructs every proposed weekly entry from the original A:V programme via `trial_source_weekday_projection.preview_source_weekdays`. For each explicitly confirmed patient identity, it compares the complete multiset of (PatientID, treatment, weekday, time, provider, source-cell) against `TrialImportPlan.entries`, rejecting missing/extra/changed sessions. This catches an important limitation of the earlier source-alignment audit: correctly pointing to C4 did not prove that the staged clock/day/provider were correctly copied. Blank weekdays + populated time expand to exactly ΔΕ, ΤΡ, ΤΕ, ΠΕ, ΠΑ. No sessions are created without time. If any F reclining or U/V afternoon/pool source fact is not mapped, or any supported cell is ambiguous, the entire patient row is held rather than imported partly.

**Integrated gate:** `trial_demo_source_projection.project_trial_demo_sources` now checks this exact source parity **before** combining PATIENTS and PROGRAM_WEEKLY_SLOTS rows. Synthetic acceptance tests added to `Tests/test_trial_patient_registry_projection.py`: source blank-day five-weekday parity, changed time under correct source-cell reference, blocked unmapped afternoon, and existing combined source/registry/weekly synthetic roundtrip. No real clinical data, native write or VBA modification was made.

**Evidence:** GitHub Actions `43758856408`, code commit `ba001d35459732724ca2c9be7f51b3864720d6d7`, completed **all four Windows/Linux synthetic and full regression CI jobs successfully**. An intermediate code-only commit before the synthetic fixture was updated had a red CI run; that intermediate result is superseded by the green `ba001d35` run. Do not rerun this accepted source gate merely for reassurance. It is not native Excel acceptance.

**Known boundary / precise next implementation:** only source-bound read-only projection and reconciliation exist, not the requested runnable XLSM. A native writer requires a fresh eligible **empty operational demo shell** and independent Excel Save/Close/Reopen, ensuring the old identities, frozen H10 artifacts and history remain untouched. Since `PATIENT_PLANNER`/MASTER is visual single-slot oriented while `PROGRAM_WEEKLY_SLOTS` supports multiple distinct day/time placements, the writer MUST expose complex weekly rows without compressing them to an incorrect single time. Confirmed private PatientIDs, clinical classification and holds/shared-provider approvals remain mandatory before importing real patient records. Target one independent practice workbook, not another analytics prototype.


## 2026-11-15 — Private clean-shell native DEMO seeding bridge

**Immediate delivery goal remains a single runnable practice workbook** for patient admission/discharge, absences, therapist replacements, new therapist/student registration and daily operations. Statistics/15-day progress remain subsequent/parallel work. Historical H2–H10 accepted proof artifacts remain frozen and are NOT the write destination.

**Implementation added**:
- `Python/tools/preflight_trial_excel_sources.py`: opt-in `allow_empty_demo_shell=True` permits ONLY an otherwise eligible macro-enabled intake containing **zero patients, zero protected history records** and exactly the `EMPTY_OR_MISSING_PATIENT_REGISTRY` readiness issue. Default preflight behavior stays unchanged. Any real patient registry, external workbook links, missing required operational sheet, malformed history or other issue continues to block.
- `Python/rehab_excel/trial_import_plan.py`: `empty_demo_shell=True` must be separately proved by preflight; `TrialImportPlan.empty_demo_shell` is not inferred and default remains False. An eligible real-source DEMO still requires independently confirmed IDs, zero review holds, zero overlaps and exact shared-provider approval.
- `Python/rehab_excel/trial_demo_native_seed.py`: guarded `seed_trial_demo_workbook` creates a unique **private output only outside a Git checkout**, from an eligible empty shell, through one isolated temp XLSM. It refuses existing output files, old patient/planner/weekly/other operational records, students, historical rows or source-hash changes. Real XLSM modifications use Excel COM, not openpyxl save. It writes PATIENTS, row-bound PATIENT_PLANNER identity formulas, granular PROGRAM_WEEKLY_SLOTS, and a full non-compressed DEMO_WEEKLY_VIEW; native Excel saves and closes before independent readback. The output is exclusively published only after verification, with no overwrite. It verifies the data read back by the actual `read_base_schedule` daily-scheduling source reader, protecting replacement-engine input against phantom/missing weekly slots. Optional mock backends are never considered proof of native Excel acceptance.
- Tests: `Tests/test_trial_demo_native_seed.py`, extensions in `Tests/test_preflight_trial_excel_sources.py` and `Tests/test_trial_import_plan.py`. These use synthetic fake VBA containers solely for structural rejects and no clinical identities.

**Explicitly NOT delivered/accepted:** no populated private real-source XLSM, no native Windows COM Save/reopen trial on the user's machine, no full UI buttons/workflow acceptance, and no patient-ID authorization. The clean shell has not been created/approved for this specific import. The legacy MASTER/PATIENT_PLANNER visual model is single-slot oriented and cannot correctly compress all multi-day/time treatment cases; DEMO_WEEKLY_VIEW preserves the complete schedule but is **not** a substitute for integrating all complex treatments into the familiar operational UI. Do not label this partial native bridge a finished practice application.

**Next true engineering gate** (not repetition of old H2–H10 tests): prepare a separately eligible EMPTY operational demo shell, bind original A:V/review/source-hashes and approved private clinical facts, perform a fresh Windows native seed on a never-existing private path, verify the workbook within Microsoft Excel and exercise admission/discharge, therapist absence and replacement, therapists/students and daily restore. Do not rerun already accepted 1,869+ synthetic history tests merely for reassurance; only new changes require scoped regression. See `Documentation/PRACTICE_DEMO_NATIVE_SEED_RUNBOOK_20261115.md`.


## 2026-11-15 — Production reader readback added to native DEMO seed

The native seed's independent post-save verification now also calls the production `read_patient_registry(stage)` parser and compares every imported PatientID, display name, room, infectious flag, status, inpatient/outpatient classification, hospital MRN and responsible doctor with the approved in-memory registry projection. This sits alongside the existing `read_base_schedule(stage)` comparison for every granular recurring treatment/day/time/provider. This closes a readback gap: matching worksheet cells alone did not prove that both production source readers could consume the native result.

Code commit: `702db942b5e9416344438f744efa7a6889dcc5e7`. This commit was made after the previously green `2b1c85bd` workflow; **do not claim the new reader-check commit has CI evidence until a run for this exact SHA completes**. No private source workbook, native Excel run, or operational UI acceptance was performed.


## 2026-11-15 — DEMO legacy planner visual mirror, no schedule compression

Delivered `Python/rehab_excel/trial_demo_planner_mirror.py` and focused `Tests/test_trial_demo_planner_mirror.py`. The private native DEMO seeder now writes a `PATIENT_PLANNER` time/day/provider for each admitted (patient, specialty) **only when** all granular weekdays share one clock and provider and no weekday repeats. Full five-day programmes display `Καθ/να`; irregular weekdays are preserved explicitly. The native readback verifies the mirror. Different clocks, providers, or same-day multiple slots remain exclusively in the full `DEMO_WEEKLY_VIEW` and authoritative `PROGRAM_WEEKLY_SLOTS`; they are NOT silently compressed into an incorrect legacy MASTER cell. The existing daily reader overlays granular weekly entries to prevent double-counting. No original clinical workbook or frozen proof artifact was modified.

**Scope:** user-visible compatibility improvement for the initial practice DEMO, not complete navigation/UX acceptance. Python CI for this new commit must be separately confirmed. Next: ensure the native template and registration/discharge/replacement VBA read the same source, then isolated Windows Excel/COM acceptance with a private empty demo shell.

## 2026-11-16 — DEMO navigation from MASTER to granular weekly appointments

Native-only DEMO installer `Python/rehab_excel/trial_demo_weekly_navigation.py` adds a dedicated **Πλήρες εβδομαδιαίο** button on `MASTER_SCHEDULE` to open `DEMO_WEEKLY_VIEW`, plus a **← MASTER** return button. It installs an isolated `modDemoWeeklyNavigation` VBA module, calls the existing `KeepApplicationShell`, and leaves the four canonical MASTER toolbar controls untouched. The private `trial_demo_native_seed.py` installs the button in its fresh staged workbook and verifies the persisted macro + both button actions via read-only Excel COM before publishing. The detailed view retains complex weekday/time/provider entries without loss; the legacy planner mirror remains intentionally limited to unambiguous single-clock/provider cases.

Synthetic test `Tests/test_trial_demo_weekly_navigation.py` checks form/module wiring, idempotent button installation and non-interference with existing MASTER navigation. The code-level GitHub CI for the navigation commit must be confirmed separately; no completed native Windows Excel source-data seed or interactive user acceptance is claimed. **Next gate:** obtain eligible private empty shell/approved identities and produce the first disposable native workbook to exercise daily absences, replacements, patient registration/discharge and student/therapist registration in one application. Do not repeat H2-H10 frozen proof gates for this source-level change.


## 2026-11-16 — DEMO patient exits, day-effective read-only enforcement

**Delivered:** `Python/rehab_excel/patient_exit_source.py` provides an optional, strict `PATIENT_EXITS` four-column sheet contract (`PatientID`, `ExitKind`, `LastActiveDate`, `Reason`). The source reader validates the exact patient identity and registered inpatient/outpatient type, unique identities/events, supported discharge vs outpatient-end actions and native Excel/ISO/Greek dates. Missing sheet is backward compatible; malformed present sheet blocks the whole read. `preview_patient_exit_from_workbook` exposes the exact recurring base-entry IDs without writing or destroying clinical history. `read_effective_schedule_for_date` returns only date-eligible recurring entries and retains the original registry and plan unchanged. An exit's last active date is inclusive. There are no Saturday/Sunday treatment sessions.

**Operational integration:** `Python/tools/prepare_daily_input_in_place.py` now calls the pure `active_entries_on` with all approved `PATIENT_EXITS` **before** materializing that day's inpatient/outpatient sessions. A patient's permanent inpatient discharge and an outpatient's program end stop future daily bookings; neither is translated into an absence, cancellation, destruction of recurring assignments, or history rewrite. Existing workbooks without this optional sheet retain legacy behaviour. The isolated `trial_demo_native_seed.py` initializes an EMPTY `PATIENT_EXITS` ledger and independently verifies it after native save, refusing a shell carrying older patient-exit events.

**Synthetic tests** `Tests/test_patient_exit_source.py`, `Tests/test_prepare_daily_input_in_place.py` and `Tests/test_trial_demo_native_seed.py` check future/past date behaviour for both patient types, schema/date/identity failures, no phantom weekend session, read-only source preservation, and isolation of old exit facts. CI for the final post-documentation commit must be checked independently.

**Acceptance boundary:** This is a working read-only source model plus a daily-input integration gate, NOT a completed operator-facing discharge button, native lifecycle event save/rollback, date-effective MASTER census view or finished combined XLSM. Do NOT advertise discharge as operationally complete until the operator can preview, confirm, save/reopen and see consistent MASTER, daily programme and history in a fresh private candidate. Do not modify frozen H2-H10 artifacts or use real patient records in public GitHub.

**Next slice:** isolated, approved native Excel lifecycle-event transaction with preview/confirmation, rollback and independent reopen; then connect discharge and outpatient end to the existing management UI and update the MASTER active census and student/therapist lifecycle controls. Do not rerun previously accepted historical proofs just for reassurance.


## 2026-11-16 — DEMO patient exit button, preview, explicit confirmation and native seed integration

**Implemented (source + synthetic CI, not yet native operator acceptance):**
- `Python/rehab_excel/trial_demo_patient_exit_vba.py` installs a separately named **Εξιτήριο / Λήξη** button on the MASTER, preserving existing MASTER toolbar and `Πλήρες εβδομαδιαίο` button. It installs `modDemoPatientExit` with `DemoPatientExitAction` only on the isolated DEMO candidate. The macro refuses non-`_DEMO_` filenames or read-only sheets, requires exactly one authorized PatientID, validates the registered inpatient/outpatient type and the existing `PATIENT_EXITS` schema/IDs, requests an exact YYYY-MM-DD *last active day* and optional reason, previews the affected recurring entry count, and requires a default-**No** explicit confirmation. Only then does it append one patient-exit event, save in Excel and recheck the in-session sheet. If writing/save/validation fails, it attempts to clear the newly appended event and save the reversal; it must still be independently reviewed after any failure. Existing PATIENTS, PATIENT_PLANNER, PROGRAM_WEEKLY_SLOTS, outpatient source and finalized history are NEVER erased by this action. It does not silently equate an exit with a daily absence.
- `Python/rehab_excel/trial_demo_native_seed.py` now installs this new UI module and verifies the macro and button action via read-only Excel COM before publishing a **new isolated** candidate. The original source and frozen H10 proof files remain untouched.
- `Python/rehab_excel/patient_exit_source.py` now also offers `verify_saved_patient_exit`, an independently reopening **read-only** verification helper: confirms exact recorded kind/date and original recurring rows still exist, but those rows disappear for that patient after the last active day. This **is not automatically invoked by the live VBA macro**, whose immediate validation is in-session only.
- Synthetic tests in `Tests/test_trial_demo_patient_exit_vba.py` and `Tests/test_patient_exit_source.py` cover button wiring, repeatable module installation, confirmation guard, rollback markers, original recurrence preservation and wrong date/type rejection.

**Next operator/native gates:** seed a private fresh eligible XLSM using approved identities and actual Windows Excel; click exit and cancel, then confirm on a separate disposable case; Save, Close, Reopen and independently read back `PATIENT_EXITS` and subsequent DAILY_INPUT for the right effective date. The present MASTER census/visuals are **not automatically refreshed** by this new VBA action, and the macro explicitly warns the operator. Do not label the exit workflow end-to-end complete until those checks, active MASTER census and menu navigation are proven. Python/Windows CI is source regression evidence, not native Excel acceptance. No patient data must be committed to GitHub.


## 2026-11-16 — DEMO effective MASTER and live patient census bridge

**Implemented on `work/afternoon-medical-scope-20261115`**:
- `Python/rehab_excel/master_projection.py`: `build_master_projection(..., target_date=date)` optionally reads the validated `PATIENT_EXITS` ledger and excludes an inpatient from the projected **active MASTER census only after** their inclusive last-active date. Without `target_date`, legacy MASTER behaviour is unchanged. Patient registry and recurring schedules/history remain immutable. The preview fails closed on invalid dates or ledger identities.
- `Python/rehab_excel/master_projection_writer.py`: optional `target_date=` on `refresh_master_projection_in_place` with standard independent readback. This creates an explicit opt-in native MASTER regeneration path, **not yet invoked from the in-session DEMO discharge action**; do not claim automatic MASTER census refresh.
- `Python/rehab_excel/trial_demo_active_census.py`: a distinct, formula-backed `DEMO_ACTIVE_CENSUS` status view linked directly to PATIENTS and PATIENT_EXITS, with its own **Ενεργοί ασθενείς** button in MASTER and **← MASTER** back action. Date cell `H2` defaults to Excel `TODAY()` and can be changed to inspect past/future dates. Every patient remains visible with status `Ενεργός` or `Εξιτήριο / Λήξη` as of H2. Last active day is inclusive. PatientID lookup coerces IDs to text so new numeric registrations can match text-stored exit events. Formula column references follow the actual PatientType column. The workbook's `RefreshDemoActiveCensus` VBA extends source-linked formulas when new PATIENTS rows are registered; a blank row-2 template exists even with zero patients.
- `trial_demo_native_seed.py` installs this view and checks its formulas, button actions and VBA module in read-only native Excel COM, in the same isolated staged candidate as the existing weekly and exit buttons. `trial_demo_patient_exit_vba.py` invokes `RefreshDemoActiveCensus` before saving a confirmed exit, so a new patient's exit is reflected without deleting anything.
- Tests: `Tests/test_trial_demo_active_census.py`, additions to `Tests/test_master_projection.py` and `Tests/test_master_projection_writer.py`, covering inclusive dates, valid/invalid requests, immutable registry, dynamic formula IDs and future admission rows. All tests use synthetic identities only.

**Important status boundary:** Python CI must be checked for the final commit. The operator-facing in-session discharge action STILL does **not** automatically rebuild the canonical MASTER display (it now refreshes the separate DEMO_ACTIVE_CENSUS). The candidate has **not** been created from private approved patient workbooks or tested by the operator in native Windows Excel; no claim of operational acceptance. Do not change canonical MASTER automatically without native Save/Close/Reopen/rollback evidence. Complete admissions/replacements/staff/student lifecycle acceptance remains pending. No patient-identifying source or workbooks go to GitHub.

**Next priority:** safe in-session native MASTER census refresh of the chosen target date (without changing historical rows), followed by one executable private DEMO and full operator flows. Previously green H2-H10 proof checkpoints are frozen; do not rerun unrelated history/stats tests. `Documentation/PRACTICE_DEMO_NATIVE_SEED_RUNBOOK_20261115.md` is the operator safety runbook.

## AUTHORITATIVE LATEST RESUME 2026-11-16 — REAL-DATA BUSINESS DEMO

**Current delivery verdict:** A working operator real-data XLSM has **not** been delivered. Prior source integrations and 1.960/1.960 Windows/Linux tests are valid code evidence, not 1:1 actual manual schedule/Excel native acceptance. Synthetic H8/CC demos are excluded from operator acceptance. No full actual clinical records have been imported into a native test in this GitHub environment. Explicitly open: F/U/V, outpatient register coverage beyond inpatient-only real A:V weekly projection, private identity/clinical approval, clean Excel shell, native COM/app UI integration, canonical MASTER date-effective refresh, save/reopen + rollback. Source manual remains production authority. New date-specific full delivery baseline and operator matrix in `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`. All historical H2–H10 proof files remain frozen.


## 2026-11-16 — F/U/V read-only gate, 1,991 tests GREEN

**Concrete new gate, not a synthetic replacement for the user's real-data DEMO:** `Python/rehab_excel/trial_special_fields_audit.py`, `Python/tools/audit_real_manual_specials.py`, `Tests/test_trial_special_fields_audit.py` provide the first **single-source-only** private read-only inspection of actual manual A:V columns **F** (reclining, known 30-minute ΦΘ-related operation), **U** (legacy afternoon, no auto-import) and **V** (morning pool). No review/clinical approval/clean XLSM required for this *diagnostic* first command. It checks original source sheet/header, SHA-256 before/after, cached and raw Excel formula values (never omit formula-only cells), missing-patient source rows, time parseability, all counts and no PII emitted. The tool **never creates sessions, guesses F/V weekdays or imports old U**. It returns 0 if F/U/V truly empty (not whole source approval), 3 for pending real special records, 2 on safety stop. Exact local command, interpretations and privacy gates: `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md`.

**Verified code evidence:** `95157af7dbb32c6cb8aca210c503ffbb7f63be17`, GitHub Actions `43804737124`, **1,991 passing pytest in both Windows and Linux**, all four full/synthetic CI jobs successful. This is not proof that the real private manual has been scanned or a working native Excel was built. **Real-data DEMO remains NOT DELIVERED.** F/V day/provider/duration-aware scheduling approval, U source scope and independently reviewed outpatient programme remain blockers. Next meaningful step is run the read-only audit against the operator-approved private CURRENT manual snapshot, capture only totals, finish explicit F/V mapping and standalone outpatient source, then eligible private native seed and full side-by-side weekday/operator acceptance. Frozen H2–H10 proof artifacts remain untouched.


## 2026-11-16 — Mandatory F/U/V source gate before real DEMO seed

`Python/tools/prepare_real_data_practice_demo.py` now calls `inspect_real_manual_special_fields(programme)` **before reading approved clinical values, building source plans or invoking native Excel**. Any populated F reclining, U legacy afternoon or V morning pool field (including Excel formula-only content with empty cache), or an orphaned source row, gives a privacy-safe **SAFETY STOP**. This ensures that even a stale review workbook cannot silently hide real original-source special therapies. F/V clock parseability is diagnostic only: **days/duration-aware scheduling and authoritative source treatment review are still required**. U remains expressly non-importable from the historical column; separately entered authorized afternoon appointments belong to their own programme. The standalone command `python Python\tools\audit_real_manual_specials.py --programme "<PRIVATE_CURRENT_MANUAL>.xlsx"` reports only aggregate counts and returns 3 for blocked real specials, 0 if no F/U/V cells are populated, 2 on structural safety failure. No original workbook or native XLSM was changed in the GitHub development environment.

**Tests:** `Tests/test_prepare_real_data_practice_demo.py` adds explicit refusal of preflight and native build before clinical approval/Excel on pending F/U/V/orphan formula cells, alongside `Tests/test_trial_special_fields_audit.py`. Verified integrated hard-stop code checkpoint `a9d2f0afd8b119b69a562475525d30d35d7cd8a3`: **1,993 tests passed on Windows and Linux, 4/4 CI jobs successful** in GitHub Actions run `43805000757`. This is source/synthetic regression evidence, not a real-manual scan or native Excel acceptance. Real private source inspection, source review, native Excel candidate and user acceptance remain **PENDING**. The new hard-stop is deliberate and must NOT be bypassed with invented patients or source omissions.

## 2026-11-16 — Existing MASTER content and formatting reused in real-data DEMO

**User clarification confirmed in source, NOT a new MASTER design:** `Python/rehab_excel/layout_contract.py` defines MASTER headers, treatment locations (D ΦΘ, E ρομποτικό, F πισίνα, G ανακλινόμενο, H εργο, I λογο, J ΕΦΑ, K ψυχολόγοι, L απογευματινό, M κατάσταση). `patient_program_master_render_vba.py` already implements `RefreshMasterPatientProgram`, `BuildTreatmentDisplay`: the actual treatment cell content is **clock (top line), weekdays (next line), optional approved provider (third line)**, wrapped and centered; `master_visual_style.py` owns existing colors, dimensions, header/identity styling, clinic banners. `PATIENT_PLANNER` has each matching `<treatment>_Ώρα`, `<treatment>_Ημέρες` and optional provider. These contracts are preserved rather than re-created.

**Integration defect discovered/targeted:** `master_projection_writer.py` must first sort admitted patients; its original one-cell formula templates, when used by the trial seeder, do NOT guarantee the already established multiline treatment display. New `Python/rehab_excel/trial_demo_master_display_bridge.py` reopens ONLY the disposable private staged `NATIVE_DEMO_STAGE.xlsm`, **requires** the already-installed `PatientProgramMasterRender.RefreshMasterPatientProgram` VBA module, calls it for each imported inpatient after the existing MASTER projection + visual style, and saves. It **does not install a duplicate formatter, change cell layout, modify the operational original or touch H2-H10 proofs**. Independent read-only `verify_private_demo_master_display` checks exact time, weekdays, optional provider and Excel WrapText for each losslessly represented (patient, treatment) in the initial trial; mismatches or a missing established VBA macro **fail closed** before publishing. The existing `DEMO_WEEKLY_VIEW` still shows all weekday/time variants where the legacy single-cell representation cannot show them losslessly. Import of actual source F/V/U special treatment fields is a SEPARATE unresolved input-admission contract; **none of the already supported treatment categories is prohibited**.

**Proof scope:** `Tests/test_trial_demo_master_display_bridge.py` checks original VBA column placements, full content/format order, sample persisted readback, deliberate tampering, missing/stale macro and seeder call order. CI for this latest commit is tracked separately. Native Windows/COM real-data candidate has **not** yet been executed. Required next actions are exact real manual source F/V/outpatient approval, private native seed and operator comparison, not another MASTER redesign.


## 2026-11-16 — Real outpatient source audit reuses existing application schema

**Current verified architecture:** Normal patient registration/program editor already writes `PATIENT_PLANNER` and uses `PatientProgramMasterRender.RefreshMasterPatientProgram` for the existing MASTER cell content. The new bulk real-data DEMO seed is a *different* initial-population path; it does not imply the old MASTER renderer was broken. Do NOT rebuild established cell formatting, controls or clinical treatment categories.

**Next separate real-data source bridge:** `Python/rehab_excel/trial_private_outpatient_audit.py` and `Python/tools/audit_real_outpatient_source.py` inspect only a PRIVATE, institution-approved current outpatient export in `.xlsx` form. They require the already-supported `PATIENTS` and `OUTPATIENT_SCHEDULE` sheets and their existing exact leading column headers. Reuse `read_patient_registry`, `read_outpatient_schedule` and `parse_day_pattern` to check outpatient identity/type/name consistency, Mon-Fri days, times, provider for ΦΘ, no robotic outpatient, per-weekday duplicate bookings, no formula-only or orphan source rows and SHA-256 read-only preservation. CLI emits only aggregate counts; never names/IDs/paths. `Tests/test_trial_private_outpatient_audit.py` uses only fictional private fixtures. **This is a diagnostic admission precondition, not an outpatient import or proof of completeness**. Full native outbound bridging from this separate approved source into the same private XLSM remains PENDING. Do not infer outpatients from the existing inpatient A:V review or omit them from the final pilot. Full command and schema: `Documentation/REAL_OUTPATIENT_SOURCE_EXISTING_SCHEMA_AUDIT_20261116.md`.

**Real workbook and operator acceptance status remains unchanged:** no current actual source files available in GitHub, no private native Excel produced or verified, outpatient source admission/merged seed and F/U/V approvals still open, 1:1 real manual comparison pending. Previous 2,002 test green code checkpoint `3cfdf61b` with GitHub run `43805735643` establishes the earlier MASTER bulk-display integration, not outpatient native acceptance. Check the new outpatient audit CI separately. Frozen H2–H10 acceptance files stay immutable.


## 2026-11-25 — Day Amendment UI/persistence preflight: BLOCKED, no positive transaction

The operator rechecked candidate `UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` (731,415 bytes; SHA-256 `1BD5E58E905F91ACB55E5A706CD8B4563B5A98C170A9EC60FE76C95617420B13`) and confirmed the same hash after inventory. The read-only `inspect_day_amendment_ui_contract.py` check failed because it did not find the expected day-save button on `DAILY_INPUT`. A subsequent `diagnose_day_save_in_session_persistence.py` run stopped before persistence because the disposable candidate failed the application artifact contract: missing VBA component `modDailyInputJournal`. No positive write, receipt, close/reopen, restore or rollback test was completed. Treat this as a candidate-specific UI/application-contract blocker, not proof that the historical H3/H5/H6 amendment workflows never worked. Earlier frozen native evidence remains scoped to its own artifacts and scenarios. Next: non-mutating inventory of actual controls and VBA components, compare with installer/inspector/application-contract source, then choose a fresh disposable candidate only after the mismatch is understood. Do not modify CC or frozen evidence files.


## 2026-11-25 — Native inventory and UI contract reconciliation

For `UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm`, read-only native inventory completed (66 VBA components, exit 0; original SHA-256 unchanged). The workbook contains `modDayAmendmentAction` (402 lines) and `modUnifiedDaySave` (70 lines), but not `modDailyInputJournal`. The UI control audit found `ekaToolbar_SaveOperationalDayMaster` on `MASTER_SCHEDULE` with `OnAction=SaveOperationalDay`, and confirmed the dispatcher exists. On `DAILY_INPUT`, only the Exit control matched; no day-save/amendment shape was found. Consequently `inspect_day_amendment_ui_contract.py` fails because its expected Day Amendment button is absent from DAILY_INPUT, while the unified save-control audit succeeds because a different save button exists on MASTER_SCHEDULE. These are distinct contracts, not contradictory results. No workbook was changed. Next: reconcile amendment button name/location and installer expectations against intended operator workflow, and separately trace the missing `modDailyInputJournal` dependency before any installation or positive transaction.



## 2026-11-25 — Amendment installer contract clarified

Source confirms that the Amendment UI contract intentionally requires `ekaToolbar_DayAmendment` on `DAILY_INPUT`, bound to `StartClosedDayAmendment`. The installer calls `install_day_amendment_action`, which replaces `modDayAmendmentAction` and adds that shape, then saves. This is distinct from the unified save button on `MASTER_SCHEDULE`, so the two inspector results are consistent. However, the Amendment module calls `DailyInputJournalApplyEnabled()`, and the CC inventory lacks `modDailyInputJournal`. The Amendment UI installer does not provision this missing journal dependency. Therefore it must NOT be run on CC yet. Next: inspect the unified-save installer and native journal provisioning/dependency contract, then plan changes only on a disposable clone.



## 2026-11-25 — Journal PREPARED bridge and Unified Save UI reconciliation

Source review confirms the native journal provisioner creates/validates only the protected journal worksheet in a disposable copy, saves/reopens it in native Excel and explicitly installs no operational hooks. The separate `daily_input_journal_vba.py` defines `modDailyInputJournal` and installs a PREPARED-only bridge whose `DailyInputJournalApplyEnabled()` deliberately returns False; full Apply/Restore is not enabled. Prepared preflight is read-only, and transition preparation remains `TRANSITION_PROPOSAL_ONLY`, with no write authorization or operational Apply/Restore.

The Unified Save installer routes from `MASTER_SCHEDULE` to `CloseOperationalDay` when there is no current final snapshot, or to `StartClosedDayAmendment` when exactly one current final snapshot exists; ambiguous state stops. It deliberately removes `ekaToolbar_DayAmendment` from `DAILY_INPUT`, because its stated UI contract keeps only Apply & Continue there and centralizes saving on MASTER. This conflicts with the standalone Amendment UI installer, which expects/creates the button on DAILY_INPUT. Reconcile these contracts before treating the missing DAILY_INPUT button as a defect or installing either component. Preserve original CC; all future assembly only on a disposable clone. Positive operational transactions remain NOT READY / NOT PROVEN.


## 2026-11-25 — DAILY_INPUT lifecycle review

Operator-provided full source output confirms the journal VBA currently supports only a PREPARED technical probe. `DailyInputJournalApplyEnabled()` deliberately returns False; the bridge validates source/handoff/payload and journal integrity, appends PREPARED, saves and independently verifies by native Excel reopen. Its failure path preserves a possibly durable journal row and fails closed; it is not business rollback.

The Python lifecycle module is a pure legal-transition model (PREPARED through COMPLETE, FAILED and ROLLED_BACK), not a persistence/orchestration engine. The phase evidence module lists native-reopen evidence obligations but explicitly registers no independent phase verifiers and does not authorize writes. Consequently no positive Apply/Restore/Amendment transaction or full-workbook rollback is yet proven. Keep the gate closed. Next: inventory authoritative handlers and persistence boundaries, implement independent phase verifiers, then native VBA orchestration, durable recovery and full rollback; fault-inject and prove on disposable synthetic workbooks before application assembly.


## 2026-11-25 — Authoritative Apply / Replacement / Restore / Amendment source inventory

Source review identifies existing handlers rather than missing functionality: Apply in `modDailyInputAction`, automatic replacement in `modGlobalReplacementAction`, manual replacement in `modReplacementOperator`, in-session ledger/replacement restore in `daily_restore_in_session_vba.py`, Amendment in `modDayAmendmentAction`, and a strict read-only full-workbook rollback package audit in `daily_input_full_rollback_audit.py`.

All operational Apply/replacement/Amendment paths reviewed are guarded by `DailyInputJournalApplyEnabled()`, currently False. Apply has local patch rollback and an outer whole-operation snapshot, but on dependent-phase failure its `IncompleteApply` branch warns of possible partial saved changes and retains the snapshot without automatically restoring it. Replacement and Amendment have local rollback attempts but explicitly do not claim independent rollback proof. The in-session Restore copies a prepared preview into the open workbook and uses an in-session recovery copy; it is not a complete DAILY_INPUT transaction rollback and its source disclaims independent Save/Close/Reopen proof. The full rollback audit compares all non-journal XLSM package members byte-for-byte and validates journal history, but requires separate trusted receipt and native Excel reopen; it is an auditor, not the executor.

Therefore it would be inaccurate to say no rollback code exists. The actual gap is integrated durable orchestration/recovery and independently proven native end-to-end transactions. Gate remains closed. Next: map each handler's mutation boundaries, rollback scope, saves, independent verifiers and existing tests; do not run positive operations or change the gate.


## 2026-11-25 — Handler/Test/Evidence/Gap coverage inventory

Reviewed test files on `work/afternoon-medical-scope-20261115`: `test_daily_input_action_vba.py`, `test_daily_input_gate2_guard_matrix.py`, `test_diagnose_daily_input_gate2_apply_worker_native.py`, `test_daily_input_full_rollback_audit.py`, `test_verify_daily_input_full_rollback.py`, `test_diagnose_daily_input_full_rollback_native.py`, `test_global_replacement_action_vba.py`, `test_replacement_operator_recovery.py`, `test_daily_restore_in_session_vba.py`, `test_diagnose_daily_input_gate2_restore_worker_native.py`, `test_day_amendment_action_vba.py`, and `test_diagnose_daily_input_gate2_day_amendment_worker_native.py`.

Coverage separates into source-contract tests, synthetic package/algorithm tests, and native Excel negative disposable-copy diagnostics. These support guards, source wiring, rejection behavior, package-audit rules and local recovery contracts. They do not establish a positive permitted Apply/Restore/Replacement/Amendment transaction through the public Excel UI with durable receipt, Save/Close/Reopen verification and independently proven rollback. No complete positive end-to-end proof was found in this reviewed set. Gate 2 remains false.

Next work package: map durable journal states and trusted receipt binding to every public entrypoint/mutation boundary; complete the transaction orchestrator and disposable native fault-injection/recovery proof before considering any positive operational test.

## 2026-11-25 — Durable phase and receipt-binding audit

The lifecycle enum and legal transitions are defined in daily_input_operation.py, but the module is pure and does not persist journal state or make Excel mutations atomic. The append-only codec in daily_input_journal.py defines OperationID | Revision | Phase | PayloadSHA256 | Payload, validates checksums/canonical payload and legal revision chains, and blocks a new operation until the previous one reaches COMPLETE or verified ROLLED_BACK. Its checksum is corruption detection, not authenticity.

The native VBA journal module currently implements only the PREPARED technical probe. DailyInputJournalApplyEnabled() is literal False. The probe validates and appends PREPARED, saves and performs independent readback; it does not implement the later phase writers or operational Apply. Apply explicitly exits while this gate is false.

daily_input_phase_evidence.py defines the native reopen proof obligations for input+ledger, projections, pending replacement inventory, replacement facts, applied archive, final workbook, failure diagnostic and full rollback. It explicitly registers no independent phase verifiers; the policy output is not evidence or authorization.

Receipt binding remains a material gap: the rollback auditor requires a snapshot digest bound by a separately verified durable transaction receipt, but its API receives expected_snapshot_sha256 from the caller and does not authenticate/retrieve that receipt. A digest alone is not an authenticity proof.

Next implementation sequence: define trusted receipt schema; implement independent native reopen verifiers for every phase; build VBA orchestration that appends phases only after verified evidence; implement full-workbook rollback executor and terminal proof; fault-inject every boundary on disposable synthetic XLSM; then reassess positive native demo transaction. Gate 2 remains false.

## 2026-11-25 — Transaction Receipt System design baseline

Added Documentation/DAILY_INPUT_TRANSACTION_RECEIPT_DESIGN_20261125.md as a design-only, non-authorizing contract. It specifies a canonical receipt envelope binding operation/source identity, service date, journal revision/phase, required evidence kind, relative artifact locator and exact artifact SHA-256, pinned verifier identity/version/method, UTC verification time, prior receipt digest and receipt digest. Receipts must match the append-only journal and phase evidence contract; workbook evidence requires independent native Excel reopen. Crash windows between artifact/receipt/journal persistence must fail closed and recover idempotently.

Trust warning: a SHA-256 chain detects corruption but cannot authenticate its creator. No trusted signing/attestation root is present in the reviewed code. Do not embed a key in VBA/Python or treat Git identity as runtime signer. Design remains non-authorizing until trust root/key custody is selected and tested. No runtime code or workbook changed; Gate 2 remains closed.

## 2026-11-25 — Structural transaction receipt codec implementation

Added Python/rehab_excel/daily_input_transaction_receipt.py and Tests/test_daily_input_transaction_receipt.py. The new pure module supports canonical JSON encoding/decoding, digest checks, strict schema and field validation, operation/phase/evidence binding, verifier ID/version registry-label matching, predecessor-chain matching, safe relative artifact locators, evidence-root containment and exact artifact byte SHA-256 verification. It is explicitly non-authorizing and cannot authenticate origin: SHA-256 is not a signature and registry labels are only trusted when supplied by a trusted build manifest.

Synthetic tests were added for round-trip, tampering, malformed fields, unsafe paths, identity/evidence/verifier drift, predecessor mismatch, artifact mutation and symlink escape. Tests have not yet been executed in this session. No runtime receipts, workbook/VBA changes or Gate 2 changes were made. Next: run focused and regression pytest suites locally, resolve failures, then map actual native verifier outputs. Trusted signing/attestation and phase orchestration remain unimplemented.

## 2026-11-25 — PREPARED receipt evidence-kind correction

After the local baseline run (68 passed, 1 skipped, exit 0), source review found the synthetic revision-0 receipt fixture used FAILURE_DIAGNOSTIC, which is semantically wrong for PREPARED. Added EvidenceKind.PREPARED_SNAPSHOT and enforced that it is exclusive to revision 0; FAILED diagnostic evidence remains reserved for FAILED. Updated the fixture and added a rejection test for phase/revision mismatch. The amended tests have not yet been executed. The prior green run applies to the prior code only. No workbook or VBA changed; Gate 2 remains closed.

## 2026-11-25 — Receipt fixture phase/revision alignment correction

Follow-up local run after PREPARED_SNAPSHOT tightening: 5 failed, 64 passed, 1 skipped. All five failures traced to receipt test fixtures using revision-zero PREPARED defaults while claiming later-phase INPUT_AND_LEDGER_REOPEN evidence, or using PREPARED_SNAPSHOT for a later revision. Production phase guard was behaving as designed. Test fixtures corrected on the authoritative branch: later-phase operation and receipt now use revision 1 INPUT_AND_LEDGER_COMMITTED, INPUT_AND_LEDGER_REOPEN and a bound PREPARED predecessor digest; missing-predecessor test explicitly removes the predecessor while retaining the correct later-phase evidence kind. Changes committed; rerun pending.

The single skip is confirmed as Windows symlink creation unavailable. Six ZipFile unraisable deallocation warnings remain unresolved and are not suppressed. No native workbook/VBA or operational handler changed; Gate 2 remains closed.

## 2026-11-25 — Receipt regression green and ZipFile resource-lifecycle follow-up

Corrected receipt regression run: 69 passed, 1 skipped, 6 warnings in 5.22s. The skipped test is the symlink-escape case because symlink creation is unavailable in the Windows environment. Six ZipFile deallocation warnings remain.

Source review found the read-only rollback journal loader requested keep_vba=True despite never writing or preserving VBA; the VBA package is independently checked by ZIP-level package inspection. Changed that loader to keep_vba=False (commit 4a7c168dfecb2eb30de79ae0dc87c3d3c4b763ec) to remove unnecessary VBA ZipFile retention. This correction is committed but NOT YET TESTED. Next local run must confirm whether warnings disappear. No workbook/VBA/operational transaction code changed; Gate 2 remains closed.

## 2026-11-25 — Remaining two ZipFile warnings isolated to synthetic non-macro fixtures

After the production read-only journal loader switched to keep_vba=False, warnings fell from six to two. Remaining warnings are attached only to formula/style variants of test_business_cell_or_package_resource_change_fails_closed. The rollback test fixture is an ordinary openpyxl Workbook saved with an .xlsm extension and contains no VBA project; two test-only load_workbook calls unnecessarily used keep_vba=True. Changed those two calls to keep_vba=False (commit 30f278bf1fe6d624fe7774ac51c3a6f3466a6378). This does not alter production XLSM behavior. The change is committed but NOT YET TESTED; rerun to confirm warning removal. Do not suppress warnings.

## 2026-11-25 — Phase-verifier contracts and report envelope added

Added daily_input_phase_verifier_contract.py and test_daily_input_phase_verifier_contract.py. This is a strict, immutable contract registry covering all ten EvidenceKind values plus canonical report creation/decoding, report digest checking, operation/revision/phase binding, transition evidence derivation from the lifecycle policy, exact contract scope and method matching, native Excel version requirements, and rollback package/PREPARED snapshot digest requirements.

Historical boundary at registry creation: PREPARED_SNAPSHOT had an existing diagnostic-only readback path but no report adapter. The later PREPARED phase-report adapter now gives it status REPORT_ADAPTER_ONLY; trusted attestation is still absent. The remaining nine entries are NOT_IMPLEMENTED. No verifier is marked IMPLEMENTED for positive phase authorization. This is not an executable native Excel verifier and does not authenticate report origin. A synthetic VERIFIED report is rejected until an independent verifier implementation exists. No caller-supplied Boolean or caller-supplied evidence list can satisfy a phase obligation; the required list is derived from the lifecycle transition. No workbook/VBA/operational handler or Gate 2 setting changed. Tests have been added but NOT RUN locally.

## 2026-11-25 — Phase-verifier contract suite execution confirmed

User executed the focused Windows suite: 110 passed, 1 skipped in 4.79s. The only skip is Tests/test_daily_input_transaction_receipt.py:208 because symlinks are unavailable on this platform. This is a confirmed green focused test baseline, not proof of native phase-verifier integration or operational Apply/Restore. PREPARED remains diagnostic-only, the other nine verifiers remain unimplemented, and Gate 2 remains closed.

## 2026-11-25 — PREPARED phase-report adapter added; execution pending

Added daily_input_prepared_phase_adapter.py and test_daily_input_prepared_phase_adapter.py. The adapter reuses the existing native PREPARED readback, captures native Excel version, binds the exact saved XLSM bytes as both evidence artifact and PREPARED snapshot digest, enforces evidence-root containment and exclusive report creation. Status is REPORT_ADAPTER_ONLY. The report is intentionally rejected by phase validation until trusted verifier attestation is integrated. No receipt/journal write, Apply/Restore enablement, workbook/VBA mutation or Gate 2 change. Tests have not yet been executed locally. Previous confirmed focused suite remains 110 passed, 1 skipped, but predates this change.

## 2026-11-25 — PREPARED adapter included in full CI regression

GitHub Actions for adapter-bearing commit `a559a35c9a618b0d83764349d3aaa48e0e9cd211`: Windows full regression 2734 passed / 71 warnings / 34.45s; Ubuntu full regression 2734 passed / 69 warnings / 30.31s. Both synthetic safety jobs passed. Warning details are suppressed by the workflow's --disable-warnings option. This validates the automated regression suite, not native COM execution against the user's real workbook. PREPARED remains REPORT_ADAPTER_ONLY pending trusted attestation. Latest code/test head `a8b1c59d5ec7faabfb8cbbce41bd80cc09582930` also passed full regression: Windows 2734 passed / 71 warnings / 35.24s; Ubuntu 2734 passed / 69 warnings / 18.87s. Both synthetic safety jobs succeeded. Later commits are documentation-only.


## 2026-11-25 — Offline internal verifier audit and next-phase preparation

A GitHub source-only review is recorded in `Documentation/DAILY_INPUT_INTERNAL_VERIFIER_REVIEW_20261125.md`. It covers all ten evidence kinds, the existing INPUT_AND_LEDGER OOXML readback, PREPARED COM readback pattern, applied archive, replacement facts, rollback audit and the report/receipt trust boundary.

Two implementation gaps must be closed before positive phase authorization: (1) transitions with multiple required evidence kinds need an exact, complete evidence bundle validator; (2) receipt binding must derive required evidence from the lifecycle transition internally instead of accepting a caller-supplied evidence tuple. Trusted verifier/build-manifest attestation is also absent; SHA-256 is not authenticity.

The next technical phase remains INPUT_AND_LEDGER native Excel COM read-only verification, layered on top of (not replacing or duplicating) the existing OOXML consistency check. The zero-change business rule must be explicitly decided before implementation. No local tests or Excel COM were run during this offline review. Latest user local run before the two warning-related commits: 109 passed, 1 symlink skip, 17 warnings. Those commits still await local rerun. Gate 2 remains closed.

## 2026-11-25 — Transition evidence bundle validator implemented and CI verified

Implemented `Python/rehab_excel/daily_input_transition_evidence.py` and `Tests/test_daily_input_transition_evidence.py`. The validator derives exact obligations from the lifecycle transition, requires one report and receipt per required EvidenceKind, rejects missing/duplicate/extra evidence, checks shared operation/revision/phase identity, cross-binds report and receipt artifact/verifier metadata, verifies receipt predecessor order and hashes exact artifact bytes.

Hardened `validate_receipt_binding()` to derive required evidence internally from a mandatory previous operation; removed caller-selected evidence lists.

Full GitHub Actions regression for `ed3bdec7e3e22a69b1162867f346321503b26498`: Windows 2740 passed / 72 warnings / 44.59s; Ubuntu 2740 passed / 68 warnings / 28.33s. Both synthetic safety jobs succeeded. Overall workflow success.

Important limitation: this is a structural completeness/consistency validator, not an authentication or authorization mechanism. The complete-bundle synthetic success test temporarily substitutes IMPLEMENTED statuses in memory only. Actual registry remains PREPARED REPORT_ADAPTER_ONLY and nine NOT_IMPLEMENTED. No trusted attestation, native Excel COM, real XLSM, journal append, Apply/Restore or Gate 2 enablement is established. Warning details are suppressed by CI and warnings remain open.


## 2026-11-25 — Trusted verifier attestation requirements and threat model

Implemented only the pre-authorization structural claim contract `daily_input_verifier_attestation.py` and synthetic fail-closed tests. Documented required separation of consistency, integrity and cryptographic authenticity in `DAILY_INPUT_TRUSTED_VERIFIER_ATTESTATION_DESIGN_20261125.md`. Signature-algorithm and claimed signer fields are checked for syntax but do not prove authenticity. `require_trusted_verifier_attestation()` is intentionally reject-only until a separately administered trust root, cryptographic signature verifier, measured build provenance, key custody and revocation/expiry policy are deployed.

No embedded key, signing facility, trust root, workbook/VBA changes, journal mutation, Apply/Restore enablement or Gate 2 changes. CI outcome will be recorded once checked. User PC is not needed for this phase.

### CI confirmation for attestation claim contract — 25/11/2026

GitHub Actions workflow `45598412259` for code head `a66088e91792b49095559bccdf9c800ede2ad7b0` completed SUCCESS:
- Windows / Python 3.11: **2756 passed**, 68 warnings, 44.39s.
- Ubuntu / Python 3.11: **2756 passed**, 68 warnings, 30.26s.
- Both synthetic safety jobs: success.

The workflow uses `--disable-warnings`; warning details are suppressed but warning totals remain visible. No native COM was run. These are tests of fail-closed metadata handling, not proof of cryptographic attestation or operational Apply/Restore. The trusted verifier signer/key custody/build pinning remain unimplemented; all claims are rejected by the current attestation boundary. Gate 2 stays closed.


## 2026-11-25 — Native INPUT_AND_LEDGER diagnostic added without operational enablement

Added a source-bound read-only native Excel COM verifier diagnostic and fake-COM tests. The verifier reuses the existing OOXML readback, independently inspects DAILY_INPUT and the protected veryHidden DAILY_CHANGE_LEDGER in a fresh macro-disabled/read-only COM instance, compares complete native business-change identities against the same preflight facts, and verifies saved-workbook hash stability. It always attempts workbook Close(SaveChanges=False) and Excel.Quit and fails closed on cleanup problems.

The INPUT_AND_LEDGER registry remains NOT_IMPLEMENTED (no report adapter, attestation, receipt/journal transition or operational authorization). Ubuntu CI on the code/test commit 6e4d9929: 2770 passed / 77 warnings / 24.10s. Windows result awaits verification. Native Excel on a real disposable XLSM remains a required PC-only test; fake COM cannot replace it. Gate 2 remains closed.

### Native INPUT_AND_LEDGER synthetic CI final confirmation — 25/11/2026

GitHub Actions run `45599055363` for code/test HEAD `6e4d99296fc6679fbaeffd4a2003a17ac165719f` finished **SUCCESS** on all four jobs:
- Windows Python 3.11 full regression: **2770 passed**, 77 warnings, 33.58s.
- Ubuntu Python 3.11 full regression: **2770 passed**, 77 warnings, 24.10s.
- Windows and Ubuntu synthetic safety jobs: both success.

Previous interim notes that the Windows outcome was pending are superseded by this final confirmed result. Warning details are suppressed by the CI workflow; do not claim zero warnings. Native COM remains mocked in these tests, not real Excel; no report adapter/trusted attestation/journal append/Apply authorization or Gate 2 activation. An independent disposable-XLSM native Excel smoke test and explicit zero-change policy remain pending.


## 2026-11-25 — INPUT_AND_LEDGER phase report adapter, non-authorizing

Added `Python/rehab_excel/daily_input_input_ledger_phase_adapter.py` and `Tests/test_daily_input_input_ledger_phase_adapter.py`. The adapter requires an exact legal PREPARED -> INPUT_AND_LEDGER_COMMITTED transition, non-empty prior validated preflight facts, a saved XLSM confined within a controlled evidence root, and an exclusive new report path in that root. It executes the existing combined OOXML/native COM read-only input-and-ledger verifier and verifies the source file SHA-256 before and after diagnostic execution. It binds the generated canonical phase report to operation, service date, revision, phase, pinned verifier contract, exact workbook evidence bytes, and native Excel version. Existing report outputs are never overwritten. Invalid phase, inconsistent result, changed workbook, outside-root path and diagnostic failure are rejected before publication.

The INPUT_AND_LEDGER contract status is now **REPORT_ADAPTER_ONLY**, like PREPARED. Eight remaining evidence kinds stay **NOT_IMPLEMENTED**. `validate_phase_verifier_report()` continues to reject INPUT_AND_LEDGER positive authorization while trusted attestation is not integrated; `require_trusted_verifier_attestation()` remains reject-only. Adapter output is neither signed nor a receipt, journal append, actual operational COM proof or Apply/Restore authorization.

Synthetic negative/positive publication tests were added. The adapter tests use an injected fake native diagnostic; they do not execute the user's Excel. GitHub CI results for the current code/test head must be recorded after completion. The prior 2770-pass cross-platform CI belongs to the preceding diagnostic-only commit and must not be misrepresented as coverage of this new adapter. No business workbook, VBA, UI, journal writer or Gate 2 change. Zero-change policy and real disposable native Excel smoke test remain open.


### Confirmed complete regression for INPUT_AND_LEDGER phase adapter — 25/11/2026

GitHub Actions run `45599851668`, code/test HEAD `de03afc8d8cc5d824c45c755d3f0861100a6b278`, completed **SUCCESS**:
- Windows Python 3.11: **2779 passed**, 79 warnings, 39.81 seconds.
- Ubuntu Python 3.11: **2779 passed**, 80 warnings, 17.06 seconds.
- Synthetic safety jobs on both Windows and Ubuntu: success.

This supersedes provisional CI-pending notes. Tests use mock/injected native diagnostic for the phase adapter. No real native Excel/COM execution on the user's machine, actual signed attestation, receipt, positive operational transaction or Gate 2 enablement. INPUT_AND_LEDGER and PREPARED are REPORT_ADAPTER_ONLY. Other eight kinds remain NOT_IMPLEMENTED. Warnings are not zero and warning details are suppressed by CI.

## 2026-11-25 — MASTER structural saved OOXML verifier (partial PROJECTIONS work)

Added `daily_input_master_projection_readback.py` to compare canonical saved MASTER formulas and patient/doctor text against an independently supplied projection; detects unexpected row/linkage/legacy tail content and workbook mutation. Corresponding synthetic tests cover expanded and standard schemas plus adversarial edits. This is not a complete PROJECTIONS_REOPEN verifier: therapist schedules, calculated results and native Excel reopen are not checked. Verifier registry remains NOT_IMPLEMENTED, and Gate 2 closed. CI outcome must be confirmed before this new module is claimed green.


### MASTER structural readback CI defect and fix — 2026-11-25

The first synthetic regression at code/test commit `0142629049c37dce62c396e55abf0175e6576d88` failed 3 new MASTER positive tests on Ubuntu (2797 passed, 3 failed, 96 warnings). The diagnosis was incorrect literal comparison against Greek text after Unicode `casefold()`: the terminal sigma in `Θάλαμος` and `Ασθενής` was normalized to `σ`, while the comparison constants had `ς`. Fixed both comparisons by casefolding the canonical header constants too in code commit `905418df1c33b4008fd2039a89b6c06d7b4ea7fd`. New Ubuntu Python 3.11 CI `45617783680` passed **2800 tests / 99 warnings / 25.97s**, and both synthetic safety jobs succeeded; full Windows result to be confirmed separately before full green. No real Excel run or workbook modification. This was a newly introduced test/code mismatch, not evidence of operational workbook correction.


### Final cross-platform regression for MASTER structural readback — 2026-11-25

GitHub Actions run `45617783680` for corrected code commit `905418df1c33b4008fd2039a89b6c06d7b4ea7fd`: **SUCCESS** overall, all four jobs successful.
- Ubuntu Python 3.11: **2800 passed, 99 warnings, 25.97s**.
- Windows Python 3.11: **2800 passed, 93 warnings, 42.87s**.
- Both synthetic safety jobs passed.

This supersedes the provisional Windows-pending status for this run. Initial three failing new tests were corrected, not ignored. Warning details are suppressed by CI, counts are retained. This is saved-OOXML verification and synthetic tests only. Full PROJECTIONS_REOPEN remains NOT_IMPLEMENTED and no native Excel, therapist projection, authenticated attestation, Apply or Gate 2 action occurred.

## 2026-11-25 — Read-only THERAPIST_DAILY distinct-clock verifier partial scope

Introduced a saved OOXML structural comparator for the two THERAPIST_DAILY provider blocks, matching the existing VBA workload strip's distinct-clock interpretation. It fails closed on unknown providers, missing or extra occupied clocks, formula/error occupancy, orphan values and source mutation; synthetic tests cover these cases. Distinct clock slots are not the same as number of sessions/patients. No guessed provider capacity, historical attendance, replacement approval or operational permission. This is not a complete PROJECTIONS_REOPEN phase verifier, not a report adapter and not real native Excel COM. Status remains NOT_IMPLEMENTED, Gate 2 closed. CI outcome pending final verification.


### Final cross-platform synthetic CI — THERAPIST_DAILY occupied clocks — 2026-11-25

GitHub Actions run `45618473467`, code/test commit `922afe7bf7e534273781f76ebdf23e600c5ef3bc`, finished **SUCCESS** across all four jobs:
- Ubuntu Python 3.11 full suite: **2815 passed, 98 warnings, 21.23s**.
- Windows Python 3.11 full suite: **2815 passed, 98 warnings, 48.36s**.
- Both Windows and Ubuntu synthetic safety jobs: success.

The warnings are nonzero and CI suppresses their detail. This validates synthetic OOXML readback and static provider/clock expectations only, not native COM, real workbook behavior, full therapist/patient session projection, trusted attestation, phase evidence completion, or Gate 2 permission. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED. No production workbook or VBA changes.


### 2026-11-25 — THERAPIST_DAILY occupied-clock boundary hardening

Review of `daily_input_therapist_projection_readback.py` found that the numeric-time decoder used `round(value * 1440) % 1440`. A valid Excel floating-point time just below 1.0 could round to 1440 minutes and silently become **00:00** instead of failing; this risks false equivalence when checking occupied clocks. Removed the modulo wrap and now reject numeric values rounding outside a 24-hour day. The preflight `expected_clocks` tuple now rejects sub-minute Python `time` values, preventing implicitly unequal resolutions from being compared to minute-based workbook occupancy. Added three synthetic adversarial tests (near-24:00 and two sub-minute expectations) in `Tests/test_daily_input_therapist_projection_readback.py`.

This is a narrowly scoped source/diagnostic hardening, not a clinical capacity-policy change, native Excel proof, authenticated verifier, nor authorization. Source code commit `ee81a24b22f1e139e8a10714f9a09245a0498790`, test commit `d6f3e17046ac208cfd2723fcfac71c6abe37f2d6`; GitHub CI result must be confirmed separately. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED; Gate 2 closed. No workbook or VBA change.


### Confirmed CI: THERAPIST_DAILY clock boundary hardening — 2026-11-25

GitHub Actions run `45619012770`, code/test head `d6f3e17046ac208cfd2723fcfac71c6abe37f2d6`, completed **SUCCESS** overall:
- Windows Python 3.11: **2818 passed, 92 warnings, 29.70s**.
- Ubuntu Python 3.11: **2818 passed, 96 warnings, 27.97s**.
- Synthetic safety jobs on both platforms: SUCCESS.

The added clock-wrap/sub-minute tests passed. This is synthetic saved-OOXML verification only; no real Excel COM/native recalculation, authenticated attestation, complete provider-session proof or production changes. Warning counts remain nonzero and details are suppressed by CI. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED; Gate 2 closed.


### 2026-11-25 — THERAPIST_DAILY assignment identity and shared-slot membership contract

Added `Python/rehab_excel/daily_input_therapist_assignment_membership.py` and `Tests/test_daily_input_therapist_assignment_membership.py`. The **pure** `verify_therapist_assignment_memberships()` comparator requires two separately sourced immutable tuples of `TherapistAssignmentFact` (preflight expected and saved-readback observed) keyed by *stable* immutable AssignmentID, each binding PatientID, treatment, normalized provider, explicit service date and exact-minute start time. It rejects missing, duplicate, swapped, stale, legacy/row-derived, cross-date or sub-minute facts even if occupied-provider-clock counts are unchanged. For same-provider simultaneous different patients, the exact group of patient IDs must match independent prior approved shared-slot membership; excess/missing/duplicate approvals fail closed. It does not invent provider limits or automatically authorize shared treatments.

**Important trust boundary:** the comparator cannot authenticate its caller or itself reconstruct actual patient assignment IDs from the display cells of THERAPIST_DAILY. Its observed tuple must be produced by a separate independently trustworthy saved-source/COM verifier in later work. The comparator is **not** a completed positive projection verifier, report adapter, trusted attestation or Apply/Restore authorization. No workbook IO, VBA mutation, journal append or Gate 2 activation. `PROJECTIONS_REOPEN` stays NOT_IMPLEMENTED. This pure unit test suite tests in-memory fictional identity sets; it cannot prove real native Excel or full clinical assignment correctness.

GitHub code/test commits `f15e366ea01c436d46c939c05709f16cc6b33cd0` / `166d7bf54345476422dad79e2c4de964891c9214`. CI outcome must be separately recorded once confirmed.


### Confirmed cross-platform assignment membership regression — 2026-11-25

GitHub Actions run `45622874406`, code/test head `166d7bf54345476422dad79e2c4de964891c9214`, finished **SUCCESS** in all four jobs:
- Windows Python 3.11: **2839 passed, 93 warnings, 60.67s**.
- Ubuntu Python 3.11: **2839 passed, 93 warnings, 19.48s**.
- Windows and Ubuntu synthetic Python safety jobs: success.

Added tests reject changed immutable AssignmentID, patient identity, treatment, provider, date, time, duplicate assignment, missing row, legacy ID, sub-minute drift, missing/extra/shared-approval mismatch. Tests are fictional in-memory examples, **not** a real workbook source or native Excel COM inspection. No patient-identifying data was copied into repository fixtures. The complete PROJECTIONS_REOPEN phase, including independent saved-source extraction/attested report, remains NOT_IMPLEMENTED and Gate 2 remains closed. CI warnings remain nonzero.

## 2026-11-25 — Read-only OOXML VBA-archive retention cleanup; local rerun pending

User's focused Windows run after the new native INPUT_AND_LEDGER and projection diagnostics: 214 passed, 1 symlink-specific skip, 40 PytestUnraisableExceptionWarning instances from `ZipFile.__del__` / `ValueError: I/O operation on closed file`. No test failures.

Source inspection found unnecessary `keep_vba=True` in read-only paths and synthetic non-VBA XLSM fixtures. Changed to `keep_vba=False` in:
- `daily_input_reader.py` (read-only reader, closes without saving)
- `daily_input_applied_state.py` (saved OOXML readback, closes without saving)
- `prepare_daily_input_journal_prepared.py` (PREPARED preflight inspection, closes without saving)
- `daily_input_master_projection_readback.py` (read-only structural verifier)
- `daily_input_therapist_projection_readback.py` (read-only structural verifier)
- the MASTER and THERAPIST_DAILY synthetic test mutation fixtures.

This does not write or save a workbook and does not alter VBA, workbook content, native COM security, transaction state or Apply/Restore behavior. The real Excel COM verifiers remain separate and unchanged. Source changes are committed; warning removal is NOT YET VERIFIED. Repeat the focused 13-file pytest command on Windows and report the actual warning count. Do not suppress warnings or infer a clean result from the code change alone.
