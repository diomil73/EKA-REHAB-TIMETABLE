# Rehab Center System


## ΕΝΑΣ ΕΠΙΣΗΜΟΣ ΟΔΗΓΟΣ ΣΥΝΕΧΕΙΑΣ, 17/11/2026

**ΠΡΩΤΗ ΑΝΑΓΝΩΣΗ:** [DIOMIL_MASTER_CONTINUITY_GUIDE.md](Documentation/DIOMIL_MASTER_CONTINUITY_GUIDE.md). Τέσσερα πλήρη κεφάλαια μέσα στο **ίδιο** αρχείο. Όταν ο χρήστης γράψει «βρες το αρχείο οδηγό» ή «ψάξε τα 4 αρχεία με οδηγίες», διάβασε αυτό, επιβεβαίωσε current branch HEAD και συνέχισε χωρίς να επαναλάβεις πράσινα τεστ ή τις σκαναρισμένες 106/764 προτάσεις. Baseline actual source 61 holds, **15** collisions (11+4), PatientIDs confirmed 0, F/U/V=5/19/0, separate outpatient + native real-data DEMO + operator acceptance PENDING. Green 2.031 tests/OS ≠ αποδεκτό real XLSM.


## 2026-11-16 — Executable private real-source DEMO preflight

**New executable interface:** `Python/tools/prepare_real_data_practice_demo.py` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Default operation is **read-only full real-source validation** for an authorized current A:V manual XLSX, confirmed review XLSX, entirely empty application-ready VBA XLSM and independently approved private `clinical_approvals.json`. Uses the strict production source alignment, all-patient approved ID check, original weekday/time/provider reconciliation, unresolved F/U/V stop gates, existing native shell checks and SHA-256 stability of the clinical-approval JSON. Every input must remain off Git checkout. No real patient name, ID, room, doctor, MRN, source path or clinical exception content is emitted to stdout/stderr on handled errors. `--build --output <PRIVATE_*_DEMO_*.xlsm>` is an *explicit* Windows + native Excel/COM, exclusive-new-file second step and repeats the same approval gates. **No real source file has been run; no real-data XLSM exists, and end-to-end practical acceptance is still pending.** Previous synthetic/H8 workbook results do not count as the operator's comparison. First actual action is an authorized private preflight with current manual and approved clinical review, not a random synthetic seed. Native failure or changed approvals after a native operation requires freezing the candidate, not treating it as accepted. See the dedicated operator CLI runbook for exact commands/JSON schema.

**Current blocker:** F/U/V and independently sourced outpatient bookings remain unavailable to the current *inpatient-only* A:V admission bridge. If occupied, the complete real-data source fails closed, never silently drops treatments. Do not alter frozen H2-H10 evidence. Update all living handoff/status/test and delivery documents again with actual source-specific counts and results **only after** the first private native execution, without publishing PII.

## ΝΕΟ ΚΡΙΤΗΡΙΟ ΑΠΟΔΟΧΗΣ 2026-11-16: REAL-DATA DEMO

Η αξιολόγηση Diomil θα γίνει σε **ιδιωτικό XLSM με πραγματικά δεδομένα ασθενών, θεραπευτών, ημερών, ωρών και θεραπειών** από πρόσφατο αντίγραφο του σημερινού χειροκίνητου προγράμματος. Synthetic/τυχαία records αποτελούν μόνο αυτοματοποιημένα fixtures, όχι DEMO αποδοχής. Παραπομπή: `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`. Το πραγματικό DEMO **δεν έχει παραδοθεί ή ελεγχθεί native**. Τα δεδομένα παραμένουν σε εξουσιοδοτημένο ιδιωτικό Windows περιβάλλον, ποτέ σε GitHub.

Σύστημα προγραμματισμού και καθημερινής λειτουργίας αποκατάστασης/φυσικοθεραπείας.

## Start here

The canonical project entry point is:

`Documentation/ONBOARDING.md`

A new developer, a new ChatGPT context, or anyone returning after a gap should read that file first. It replaces the old scattered onboarding/recovery reading order and explains the product, architecture, locked business rules, UX direction, safety rules, current state and where to continue.

## Product direction

The `.xlsm` workbook is the execution/storage host, but the normal user experience is intended to behave like dedicated software rather than a classic Excel workbook. Users work through the MASTER screen, top-level buttons, forms and guided workflows. Raw worksheets, row numbers and internal IDs are not normal user-facing concepts.

Architecture split:

- **Excel / VBA**: application shell, forms, presentation and native workbook interaction.
- **Python**: scheduling/business logic, availability, conflicts, replacements, workload, validation and optimization.
- **Tests**: executable contract for business rules and regressions.

Development order remains:

1. define the rule,
2. write the test,
3. implement in Python,
4. make tests green,
5. then integrate with Excel.

## Baseline and write safety

- `Rehab_Center_System_v27_1.xlsm` is immutable baseline/reference.
- Never mutate the baseline directly during development.
- Do not save production/staging `.xlsm` files through `openpyxl`.
- Real workbook writes use native Microsoft Excel/COM on Windows after the required safety checks.

For current status, exact resume point and implementation ownership, follow the document map inside `Documentation/ONBOARDING.md`.


## 2026-11-15 — Project-wide delivery register

The consolidated, evidence-based project delivery report is `Documentation/PROJECT_DELIVERY_MASTER_REPORT_20261115.md`. It brings together the H2–H10 delivery evidence, Recovery track, outpatient/registration/recurring-change workstreams, real-source admission, afternoon scope, protection gates and remaining release blockers. It does not replace specialist proofs or change any frozen artifact. Use `Documentation/CURRENT_HANDOFF.md` for the exact next operator action.

 
## 2026-11-25 — Latest DAILY_INPUT verifier checkpoint

For the current verifier continuation, use `Documentation/CURRENT_HANDOFF.md` and `Documentation/DIOMIL_MASTER_CONTINUITY_GUIDE.md`. The partial native MASTER identity diagnostic, its strict scope and next implementation steps are documented in `Documentation/DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md`. It does not complete PROJECTIONS_REOPEN; Gate 2 remains closed.

 
The latest DAILY_INPUT projection diagnostics and their strict non-authorizing boundaries are recorded in `Documentation/DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md` §7–§8. The native MASTER and THERAPIST_DAILY readback primitives are fake-COM tested; the complete expected therapist grid builder, AssignmentID membership proof, trusted attestation and real Excel acceptance remain pending.

 
The current continuation has advanced through the full 126-cell expected THERAPIST_DAILY grid builder. Its latest cross-platform synthetic regression is GitHub Actions run `45660810319` (2866 passed on Windows and Ubuntu). The expected-grid builder still consumes independently validated source facts from its caller; it is not yet a full source/ledger verifier or PROJECTIONS_REOPEN phase adapter. See `Documentation/DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md` §7–§9.
