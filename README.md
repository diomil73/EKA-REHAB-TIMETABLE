# Rehab Center System

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
