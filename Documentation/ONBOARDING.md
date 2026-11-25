# Project Onboarding


## ΕΝΑΣ ΕΠΙΣΗΜΟΣ ΟΔΗΓΟΣ ΣΥΝΕΧΕΙΑΣ, 17/11/2026

**ΠΡΩΤΗ ΑΝΑΓΝΩΣΗ:** [DIOMIL_MASTER_CONTINUITY_GUIDE.md](DIOMIL_MASTER_CONTINUITY_GUIDE.md). Τέσσερα πλήρη κεφάλαια μέσα στο **ίδιο** αρχείο. Όταν ο χρήστης γράψει «βρες το αρχείο οδηγό» ή «ψάξε τα 4 αρχεία με οδηγίες», διάβασε αυτό, επιβεβαίωσε current branch HEAD και συνέχισε χωρίς να επαναλάβεις πράσινα τεστ ή τις σκαναρισμένες 106/764 προτάσεις. Baseline actual source 61 holds, **15** collisions (11+4), PatientIDs confirmed 0, F/U/V=5/19/0, separate outpatient + native real-data DEMO + operator acceptance PENDING. Green 2.031 tests/OS ≠ αποδεκτό real XLSM.


## Active real-source files actually inspected — 2026-11-17

The two uploaded REAL XLSX attachments have now been independently opened **read-only**. The original file's actual name has extra spaces: `ΠΡΟΓΡΑΜΜΑΤΑ  ΘΕΡΑΠΕΙΩΝ ΑΣΘΕΝΩΝ 10-11-26 .xlsx`; partner identity proposal is `DIOMIL_ΕΛΕΓΧΟΣ_ΤΑΥΤΟΤΗΤΩΝ_ΠΡΟΤΑΣΕΙΣ_20261112.xlsx`. Content-verified 106 matching candidate source patients and 764 weekday entries, but 0 confirmed PatientIDs, F/U blockers 5/19, 61 holds, 11 overlaps, and a separate noncanonical 10-row auxiliary outpatient sheet. The originals were not changed. **This is a private real-source diagnostic success, NOT an authorization to write a demo or clinical production workbook.** Read the new detailed privacy-safe `Documentation/REAL_SOURCE_UPLOADED_PAIR_READONLY_AUDIT_20261117.md` before continuing. Current/private approvals, outpatient conversion and native Excel/COM acceptance remain pending.


## Filename-level real source identification — 2026-11-16

The operator has identified `ΠΡΟΓΡΑΜΜΑΤΑ ΘΕΡΑΠΕΙΩΝ ΑΣΘΕΝΩΝ 10-11-26.xlsx` (original manual A:V snapshot, 10 Nov 2026) and `DIOMIL_ΕΛΕΓΧΟΣ_ΤΑΥΤΟΤΗΤΩΝ_ΠΡΟΤΑΣΕΙΣ_20261112.xlsx` (identity suggestions from 12 Nov 2026). These names are now the diagnostic locator, **not proof of access to their bytes, current freshness or independent PatientID approval**. Read `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md` for the exact non-identifying, read-only Windows commands. No original clinical workbook was imported or altered in the GitHub-only environment. All established MASTER/VBA rendering remains unchanged, and native/real-data/operator acceptance is still pending.


## Active real-data practice demo code gate — 2026-11-16

Use `work/afternoon-medical-scope-20261115`, not the October branch. The existing in-session VBA MASTER treatment renderer stays unchanged. The next build consumes two separately approved real sources: the fresh inpatient A:V original plus its confirmed review/clinical classification, and a *separate* current outpatient export with the existing PATIENTS and OUTPATIENT_SCHEDULE columns. Code is now joined and regression GREEN at `70e1b88f6e3289bbc79fe411d4724e327622ff7f` ([Actions 43809724956](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43809724956), 2,029 tests each on Windows/Linux). The operator CLI requires `--outpatients` on full `--build`; read-only checks can be run before all files are approved. Read `CURRENT_HANDOFF`, `REAL_DATA_DEMO_OPERATOR_CLI_20261116`, `REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116`, and `REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116`.

No patient-identifying source/output may be committed. The native actual-data XLSM and end-to-end real operator acceptance remain PENDING. F/U/V source holds must never be bypassed. Protect the original manual, the immutable baseline and all frozen H2-H10 proof artifacts.


## Current practice-DEMO orientation — 2026-11-16

For the latest user goal (one real-data evaluation DEMO, **not** a synthetic showcase), work on `work/afternoon-medical-scope-20261115`, not the older `feature/outpatient-scheduling`. Read `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`, `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`, `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md` and `Documentation/REAL_OUTPATIENT_SOURCE_EXISTING_SCHEMA_AUDIT_20261116.md` alongside this onboarding and CURRENT_HANDOFF. Direct GitHub verification at `b276f7b430c42cf1df46d791a122bfd2606dcf8f`: CI run 43806683353, 2,021 Windows tests + 2,021 Linux tests, 4/4 jobs GREEN. These are not real native/operator results.

Preserve the existing daily patient-program MASTER VBA display. Bulk initial import is a different entry path and must preserve its semantics without redesigning correct production behaviour. Current F/U/V holds, independent outpatient source and source-approved private patient data are release gates. Existing GitHub-tracked baseline is immutable, not a substitute for a clean empty native DEMO shell. Do not mistake code tests or source audits for real-data native acceptance.


## 2026-11-16 — Executable private real-source DEMO preflight

**New executable interface:** `Python/tools/prepare_real_data_practice_demo.py` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Default operation is **read-only full real-source validation** for an authorized current A:V manual XLSX, confirmed review XLSX, entirely empty application-ready VBA XLSM and independently approved private `clinical_approvals.json`. Uses the strict production source alignment, all-patient approved ID check, original weekday/time/provider reconciliation, unresolved F/U/V stop gates, existing native shell checks and SHA-256 stability of the clinical-approval JSON. Every input must remain off Git checkout. No real patient name, ID, room, doctor, MRN, source path or clinical exception content is emitted to stdout/stderr on handled errors. `--build --output <PRIVATE_*_DEMO_*.xlsm>` is an *explicit* Windows + native Excel/COM, exclusive-new-file second step and repeats the same approval gates. **No real source file has been run; no real-data XLSM exists, and end-to-end practical acceptance is still pending.** Previous synthetic/H8 workbook results do not count as the operator's comparison. First actual action is an authorized private preflight with current manual and approved clinical review, not a random synthetic seed. Native failure or changed approvals after a native operation requires freezing the candidate, not treating it as accepted. See the dedicated operator CLI runbook for exact commands/JSON schema.

**Current blocker:** F/U/V and independently sourced outpatient bookings remain unavailable to the current *inpatient-only* A:V admission bridge. If occupied, the complete real-data source fails closed, never silently drops treatments. Do not alter frozen H2-H10 evidence. Update all living handoff/status/test and delivery documents again with actual source-specific counts and results **only after** the first private native execution, without publishing PII.

## ΝΕΟ ΚΡΙΤΗΡΙΟ ΑΠΟΔΟΧΗΣ 2026-11-16: REAL-DATA DEMO

**Διάβασε πρώτο:** `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`. Η τελική εργασία χρήστη ΔΕΝ είναι synthetic-only: απαιτεί ιδιωτικό, ενημερωμένο **πραγματικό** πρόγραμμα ασθενών/θεραπειών/ωρών/θεραπευτών, για πλήρη σύγκριση με το σημερινό χειροκίνητο. Αντιμετώπισε τυχόν παλαιές αναφορές για «synthetic-only pilot» ως ιστορικό scope που έχει **υπερκεραστεί** για την επιχειρησιακή αξιολόγηση. Συνθετικά σενάρια επιτρέπονται μόνο ως CI fixtures. Το live real-data Excel δεν υπάρχει ακόμη και δεν είναι επιτρεπτό να αποθηκευτούν προσωπικά δεδομένα στο GitHub.

**Προτεραιότητες:** A:V/review/clinical approvals με SHA, F/U/V/holds χωρίς παράλειψη, νέα private native XLSM με πραγματική πλήρη σύγκριση ΔΕ–ΠΑ, καθημερινές ενέργειες και save/reopen, μετά analytics/appearance. Code checkpoint 1.960 GREEN Windows/Linux run 43769366651, **όχι native proof**.

Last updated: 2026-11-12

## Purpose of this file

This is the **single starting point** for a new developer, a new ChatGPT context, or anyone returning to the project after a long gap.

The old approach of discovering the project by reading several unrelated files in an uncertain order is retired. Start here first. The specialist documents listed below remain authoritative for their own domain, but this file explains what the system is, how the pieces fit together, what rules must not regress, and where to continue.

## What this product is

The project is a rehabilitation / physiotherapy scheduling and daily-operations system built on an Excel `.xlsm` workbook, but the final user experience is intentionally **application-like software**, not a normal Excel workbook.

The user should work through the MASTER screen, top-level buttons, forms and guided workflows. Worksheet tabs, raw grid navigation, internal sheet names, row numbers, internal IDs and backend mechanics are not part of the normal clinical workflow.

Excel remains the execution/storage host and the place where presentation, forms and some reporting live. Python increasingly owns scheduling/business logic because it is testable and deterministic. VBA remains only where native Excel interaction is needed and is retired incrementally rather than through a risky all-at-once rewrite.

## Canonical architecture

- **Excel / VBA**: application shell, forms, buttons, native workbook interaction, presentation and selected in-session writes.
- **Python `rehab_core`**: scheduling rules, availability, conflicts, replacements, workload, daily state, validation and optimization logic.
- **Python `rehab_excel`**: workbook adapters, render plans, bridges, VBA generators/installers and Excel-specific integration.
- **Python `tools`**: developer/verification CLIs and smoke builders.
- **Tests**: business-rule and regression contract. Development order is rule -> test -> Python implementation -> green tests -> Excel integration.

Do not move business logic into VBA merely because the UI is Excel-based.

## Workbook and source safety

- `Rehab_Center_System_v27_1.xlsm` is the immutable baseline/reference workbook.
- Development never mutates the baseline directly.
- Real workbook writes use native Microsoft Excel/COM on Windows after safety checks.
- Do not save production/staging `.xlsm` files through `openpyxl`.
- Preview/smoke workbooks are disposable copies; verify the trusted source hash remains unchanged.
- Preserve VBA and narrow write zones.
- Windows tests are always run with `python -m pytest`, never bare `pytest`.

## Authoritative scheduling model

Recurring/base schedule and daily facts are different layers.

Daily operational truth is reconstructed from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Daily absences, cancellations, no-shows, suspensions and replacements are overlays. They must not destructively rewrite the recurring/base programme.

Locked daily recomputation order:

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacement facts,
6. resolve remaining orphans by valid placement or `Αναστολή θεραπείας`,
7. only then expose capacity to new patients.

Operational capacity is timeslot-based. Existing displaced patients have priority over new patients. `MASTER_SCHEDULE` and `THERAPIST_DAILY` must represent the same effective daily state.

Every concrete-date orphan must resolve to exactly one of:

- a valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

A day is not finalized while unresolved orphans remain.

## Replacement planning: current contract and open optimizer work

Replacement lifecycle/reconstruction correctness is already established, but replacement **assignment quality** is not considered complete.

When multiple therapist absences are known for the same date, all absences must be entered first and all displaced sessions must be considered together. Do not finalize replacements greedily one therapist or one patient at a time.

The future global solver must:

- maximize valid placements and minimize unresolved patients,
- protect scarce options globally,
- never let a flexible patient consume another patient's unique viable slot when a complete solution exists,
- prefer the original time unless that harms a higher-priority constraint,
- keep ΦΘ and reclining treatment adjacent/nearby where feasible and reduce unnecessary waiting,
- keep provider workload balanced; about one timeslot difference may be unavoidable, while a two-timeslot imbalance should be avoided when feasible,
- respect patient absences, existing treatments, provider availability/capability, student limits and all other hard constraints,
- treat pool and robotic operating hours as persistent configuration once implemented,
- treat 13:00 as a weak soft preference to avoid, never a forbidden time,
- use double-booking/over-capacity only as a last-resort explicit user-approved exception.

Canonical scarcity example: if patient A can use 09:00, 10:00 or 11:00 and patient B can use only 09:00, 09:00 is protected for B when a complete feasible assignment exists.

Some clinical/operational knowledge cannot reasonably become mandatory data entry, such as therapist/patient physical suitability, personality compatibility or a frail patient's tolerance for waiting. These remain selective constraints/notes or explicit user judgment. The optimizer must not invent them.

## Outpatients

Outpatients are operationally real and count normally in workload, capacity, recurring schedules, daily therapist schedules and replacement calculations.

They do not belong in the inpatient planner UI. Their recurring schedule is stored separately in authoritative `OUTPATIENT_SCHEDULE` data and merged into the common scheduling engine.

Important UI rule: `OUTPATIENT_SCHEDULE` is a hidden/protected storage sheet, **not** an end-user workspace. The user must not be asked for worksheet row numbers or internal IDs such as `outpatient:<row>`.

Create/edit should be exposed through an application-style management flow that shows business data such as patient, treatment, time, days and therapist and carries the internal target invisibly.

Outpatient visual rule in `THERAPIST_DAILY`:

- light blue only,
- never infectious/yellow,
- never robotic,
- do not inherit another patient's cell colour.

## Application UX do's and don'ts

### Do

- expose normal workflows through the MASTER/top menu, buttons, forms and guided screens,
- show business concepts, not storage mechanics,
- keep internal IDs, sheet names and row numbers invisible,
- keep technical/data sheets hidden or protected,
- make create/edit/replacement/conflict-resolution flows usable without the Excel ribbon, worksheet tabs or raw grid,
- make success messages useful: tell the user what happened and what to do or check next.

### Don't

- require worksheet navigation for ordinary work,
- ask the user to type row numbers,
- expose internal English lifecycle strings or backend implementation details,
- send the user to a hidden data sheet merely to select a record,
- design UI that only works if native Excel chrome is visible,
- turn developer smoke/debug scaffolding into the final workflow.

## Current state at 2026-10-20

Closed/green:

- restart-safe authoritative daily reconstruction,
- replacement fact lifecycle and restore correctness,
- per-row / selected / restore-all native restore correctness,
- exact daily statistics reconstruction,
- combined outpatient recurring smoke installation,
- outpatient recurring CREATE native flow.

Deferred:

- further restore performance tuning. Current low-to-mid 20-second end-user latency is accepted for now unless it regresses.

Active:

- replace the developer-only outpatient edit selector that asks for an `OUTPATIENT_SCHEDULE` row number with an application-style schedule-management/selection UI,
- then native-test same-entry edit and conflict handling,
- later implement the global multi-absence replacement optimizer described above.

## Exact resume path

For the current branch, after reading this file:

1. Read `Documentation/CURRENT_HANDOFF.md` for the exact current checkpoint and next commands/scenario.
2. Read `Documentation/DECISIONS.md` before changing business rules or UX semantics.
3. Use `Documentation/IMPLEMENTATION_MAP.md` to locate ownership in code.
4. Use `Documentation/PROJECT_STATUS.md` for milestone/risk context.
5. Use relevant focused tests as executable specification.

Do not restart work from old chat history or old smoke assumptions when these repository documents disagree.

## Documentation roles

The previous multi-file onboarding/recovery order is replaced by this file as the single front door. The specialist files are **not deleted** because they serve different living-source roles:

- `Documentation/ONBOARDING.md`: first read, system overview, non-negotiables, reading map.
- `Documentation/CURRENT_HANDOFF.md`: exact resume point, latest native/test checkpoint and immediate next steps.
- `Documentation/DECISIONS.md`: durable business, architecture and UX decisions.
- `Documentation/PROJECT_STATUS.md`: milestone state, risks, open/deferred work.
- `Documentation/IMPLEMENTATION_MAP.md`: implementation ownership and code locations.
- `CHANGELOG.md`: release/history of implementation-relevant changes.
- `UI_APPLICATION_ROADMAP.md`: detailed product/UI roadmap and historical design specification when deeper product context is needed.
- Other `Documentation/*.md`: focused contracts for a specific subsystem; consult only when working on that subsystem.

This means a newcomer no longer needs to read seven large files in sequence before understanding the project. Read **this file first**, then only the specialist document needed for the task.

## Documentation maintenance rule

Documentation is part of implementation. In the same work cycle as a meaningful change, update the appropriate specialist source and, when the change affects the project-wide mental model, update this onboarding file too.

Chat history is never the sole source of truth.


## Current milestone update — 2026-11-04

H7 patient/therapist history and statistics is now source-green and user-confirmed in fresh disposable BF. The integrated toolbar, retained RefreshAll action, both report surfaces and direct previous-screen navigation passed native manual use. See `Documentation/CHECKPOINT_H7_BF_MANUAL_ACCEPTANCE_AND_PROTECTION_NEXT_2026-11-04.md`.

The active hardening milestone is H7 report-protection correction followed by worksheet edit protection. The combined 33-test source gate is now green (33 passed in 4.13s); the next authorized native candidate is a fresh BI from trusted AZ, followed by a complete read-only audit before manual acceptance. Do not blanket-lock sheets until sanctioned input ranges and VBA/Python write paths are mapped. BF is frozen green smoke evidence, not production. Exact resume commands live in `Documentation/CURRENT_HANDOFF.md` and `Documentation/SESSION_RECOVERY_H7_PROTECTION_BG_FAILURE_2026-11-04.md`.


## Current exact recovery pointer — 2026-11-04

The latest active work is the integrated H7 report-protection correction after BG and BH failed native smokes. BF is accepted/frozen; AZ is the trusted source; BG and BH are frozen failed evidence. The patient report installer regression has been corrected in source and the combined 33-test gate is green (33 passed in 4.13s). The next authorized native candidate is BI from AZ, followed by a complete read-only audit before any manual use. For exact commands, commit context, artifact handling and read order, start with:
`Documentation/SESSION_RECOVERY_H7_PROTECTION_BG_FAILURE_2026-11-04.md`
then read `Documentation/CURRENT_HANDOFF.md`. Do not rely on chat history to reconstruct this state.

## Latest native proof — BI audit GREEN; manual verification pending (2026-11-07)

The operator has completed the 33-test source gate (33 passed in 4.13s), fresh BI creation from frozen AZ, all H7 installers and their persisted read-only reopen checks, and a complete separate BI worksheet-protection audit. BI now persists **both** PATIENT_HISTORY_REPORT and THERAPIST_HISTORY_REPORT protected, with drawing objects unprotected to retain UI button access; four technical history sheets remain protected/veryHidden, remaining 17 sheets were not blanket-protected, and audit reports ReadOnly=True / Workbook unchanged=True. This resolves BH's patient-report protection regression at the persisted-state level. **Manual button/report-mode verification on BI has NOT yet been reported.** Do not mark BI fully accepted or broaden protection until that manual smoke is green. Start from the latest appended section of Documentation/SESSION_RECOVERY_H7_PROTECTION_BG_FAILURE_2026-11-04.md and Documentation/CURRENT_HANDOFF.md; older pending-build instructions above are historical.

## Updated immediate work — 2026-11-07: H8 analytics UX before broad protection

BI passed its source/install/protection audit, but operator manual review did not accept the statistics screens: patient numeric IDs instead of names, no clear participation percentage, therapist workload/absences insufficiently exposed, period coverage not explained, no overview/individual charts, and RefreshAll asks for repeated OK confirmations. BI is **not product-accepted** and must not be overwritten. The next task is source-level H8 analytics UX and KPI definition, not protection of the remaining worksheets. The canonical requirements and current investigation are in `Documentation/H8_STATISTICS_DASHBOARD_OPERATOR_FEEDBACK_2026-11-07.md`, and the newest section of `Documentation/CURRENT_HANDOFF.md` is authoritative.


## AUTHORITATIVE LATEST STATE — 2026-11-12 (supersedes old "current" pointers above)

**Read first:** `Documentation/AUTHORITATIVE_RESUME_20261112.md`, then the latest appendix in `Documentation/CURRENT_HANDOFF.md`. Earlier 2026-11-04/07 H7 and 2026-11-08 H10 "next" commands above are historical, not the next operator instructions.

- Real Microsoft Excel (operator-confirmed): isolated CC read-only open/close, synthetic new-sheet Save/reopen, and one-shot **179 immutable recurring AssignmentIDs** persisted in `UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm`. All three proof candidates are frozen. Underlying synthetic CC and production MASTER unchanged.
- Last fully checked cross-platform Python CI: `43313086923`, **1,646/1,646** on Windows and Linux. GitHub COM mocks are not native Excel proof.
- New optional `RECURRING_CHANGE_EVENTS` source reader handles effective-dated changes, immutable IDs, chronology, duplicates and historical projection. It is **READ-ONLY**; no native event write, permanent change UI, rollback or production transfer is authorized.
- Real-data demo MASTER remains **blocked** by unconfirmed patient identities, source holds and patient/provider time conflicts. Never confuse synthetic acceptance with actual patient admission.
- Immediate development path: synthetic-only event append plan -> new exclusive-copy native append/reopen proof -> rollback/compensation -> UI integration -> separate clinical gate. Do not rerun frozen candidate A or publish a clinical file.
- Documentation contract: update ONBOARDING, CURRENT_HANDOFF, PROJECT_STATUS, IMPLEMENTATION_MAP, DECISIONS, CHANGELOG and each relevant runbook with tested status, exact commits, safe runnable commands, failure contingencies and frozen-file ledger in the same cycle. User approval is not required for ordinary source/documentation work; real Excel operation must be run locally.


Latest verified append-plan CI: `43314844755`, 1,652 full Python regressions passed on each of Windows and Linux (not real native Excel). The new source-only event append plan is documented in `Documentation/AUTHORITATIVE_RESUME_20261112.md`. No operator action is required until a new synthetic native runbook exists.


### 2026-11-13 real-source demo preflight preparation (documentation-only)
Read `Documentation/REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md` for the latest private-input preflight, safe single-block PowerShell, stop rules, precise unsupported import boundary and weekend native acceptance plan. Tracked `trial_admission_gate.py`, `preflight_trial_excel_sources.py` and `audit_pilot_source_readiness.py` were inspected, but **no code, live private source, Windows Excel, or full tests were executed in this checkpoint**. No real-data XLSM was produced. The read-only gate is NOT an importer; strict gate still blocks on unresolved holds, overlaps and shared slots. Next work requires confirmed local private source revision and separate eligible pilot intake, then actual supported native importer engineering and one-shot proof. Do not reuse synthetic CC or frozen artifacts; keep clinical values off public GitHub.


### 2026-11-13 source-only trial import planning checkpoint
Added `Python/rehab_excel/trial_import_plan.py` (commit `b0e9d48`) and `Tests/test_trial_import_plan.py` (commit `0a9bc4c`). The plan reuses strict read-only combined preflight and exact approved `SourceWeeklyEntry` admission, cross-checks counts, binds both source SHA256 values, rejects changed sources, and hides private entries from its object repr. **It does not write an XLSM, implement a native importer, or authorize clinical use.** The new test file has six synthetic cases, but **tests/CI have NOT yet been observed passing**. Before native integration run `python -m pytest Tests/test_trial_import_plan.py Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q` in one PowerShell fail-fast block; stop on red. No private patient data uploaded. Current real staging remains unverified. More work is needed to build the new isolated native importer and independent readback.


### 2026-11-13 verified-schema weekly projection SOURCE checkpoint
Source-only modules added: `Python/rehab_excel/trial_weekly_projection.py` (`f90d4b6`) and seven synthetic tests `Tests/test_trial_weekly_projection.py` (`f2e6268`). This pure bridge maps independently approved `TrialImportPlan` entries into existing `PROGRAM_WEEKLY_SLOTS` six-column headers, with strict inpatient ID membership, no silent omissions, treatment/provider checks and patient-time collision rejection. The synthetic roundtrip test invokes the **existing `read_weekly_slots_sheet`** against an in-memory workbook. No real patient data, Excel file output, COM write, native acceptance or approved clinical import yet. New tests are **not confirmed green** unless a subsequent CI/native result explicitly proves that. Run `python -m pytest Tests/test_trial_import_plan.py Tests/test_trial_weekly_projection.py Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q` as one fail-fast PowerShell gate. Still missing: approved patient registry detail mapping, a real source-bound native writer, independent readback, privacy and operator proofs. Do not treat `PROGRAM_WEEKLY_SLOTS` alone as a complete operational demo.


## 2026-11-15 — Evidence-based baseline matrix and remaining scope clarification

Created `Documentation/INTEGRATED_TEST_BASELINE_MATRIX_20261115.md` after reviewing the H2-H8 acceptance checkpoints, H8 dashboard feedback, H10 therapist chart requirements, `STATISTICS_HISTORY_DESIGN.md`, `AUTHORITATIVE_RESUME_20261112.md`, and the current afternoon branch handoff.

Key distinction: most core daily mechanisms have substantial operator-confirmed native evidence (daily absence Apply/Restore, day close/amendment, therapist and patient session history, contextual day save, integrated reports). H8 dashboard artifact BL is manually accepted for its specifically recorded UI scope. This does **not** mean every analytics feature is complete. The patient 15-day progress/evaluation workflow remains design-only. H10 therapist charts/management have partial native/source evidence and later visual follow-ups; genuine therapist absence/working-day history is unavailable. Other explicitly open mechanisms remain global multi-absence optimization, permanent recurring-change native write/rollback/UI, real-source native importer/admission, and final worksheet protection classification. Afternoon read-only UI is source-green at 49/9.22s, but has no native Excel acceptance or save path.

The matrix records exact accepted/frozen artifact names and scope boundaries, and is the guide for selecting component references when building the single full-application synthetic test workbook. Frozen proof files are evidence only, never writable build targets. No workbook was opened or modified in this documentation review.

Read `Documentation/INTEGRATED_TEST_BASELINE_MATRIX_20261115.md` for the exact native-accepted preview artifact per workflow, the scope each one proves, the statistics/15-day progress boundary, and remaining mechanism gaps. Do not select a workbook by filename alone.


## 2026-11-15 — Full four-candidate structural comparison; three distinct VBA projects

The complete uploaded report `DIOMIL_H10_STRUCTURE_20261115_165132.txt` was inspected, superseding the earlier interim note that only the final candidate tail had been received.

### Structural comparison

| Candidate | Whole-file SHA-256 | Size | Sheets | VBA bytes | VBA project SHA-256 | Distinctive sheet |
|---|---|---:|---:|---:|---|---|
| `UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm` | `4df997b6a5e5b0b56b37127ee00dd8710cc26c7c118f1742d58c9969bcbe7afa` | 629,797 | 26 | 920,576 | `8613575ee50826394f336e4f2fd864579f087a2b1e2da4fc9883582e54a1e800` | `RECURRING_ASSIGNMENT_IDS` A1:C180, 540 populated cells |
| `UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm` | `fa9765da7079137bcb1ff935c3d08bc84b409c8335f9cd719f3583b89e4fc67d` | 622,070 | 26 | 921,088 | `7742d872c953f2b026f66097f5a61df84ef18ac151b523e88e901f10bc5a0193` | `DIOMIL_NATIVE_PROBE` A1 only |
| `UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm` | `1bd5e58e905f91acb55e5a706cd8b4563b5a98c170a9ec60fe76c95617420b13` | 731,415 | 25 | 1,305,600 | `285b7f4aa20a195535d300928f26b2884255797c1817d4f2a7e32ac82b4033b9` | no extra sheet |
| `UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` | **identical to preceding row** | 731,415 | 25 | 1,305,600 | **identical to preceding row** | byte-identical duplicate |

There are therefore **three distinct workbook/VBA fingerprints**, not four. The third and fourth candidates have identical whole-file SHA-256, not just identical VBA projects. They are byte-for-byte duplicates.

All four report 5 defined names, 11 drawings, 2 charts and 2 VML drawings; no external links, connections, query tables, pivot tables/caches, slicers, ActiveX, embeddings or custom UI were reported. The three structures share the same core 25-sheet application layout. Candidate 1 adds the recurring-assignment identity ledger; candidate 2 instead adds a one-cell native diagnostic probe.

Common formula/cache metrics:
- `PATIENT_PLANNER`: 2,745 formulas / 465 cached results.
- `MASTER_SCHEDULE`: 1,021 / 549.
- `STATISTICS`: 43 / 43.
- `PATIENTS`: 549 / 93.
- `OUTPATIENT_SCHEDULE` is header-only; `PATIENT_HISTORY_REPORT` is empty.
- `PATIENT_SESSION_HISTORY` has 523 rows; `THERAPIST_DAY_HISTORY` and `THERAPIST_ATTENDANCE_HISTORY` each have 151 rows; `FINALIZED_DAY_SNAPSHOTS` and `DAY_HISTORY_EVENTS` each have 26 rows.

Low cached-result counts are not by themselves proof of formula defects; Excel recalculation/display behavior requires a separate controlled native check. `_PY_LISTS` is a heuristic role-hint false positive unless separately shown to contain patient/staff rows.

The uploaded report ends with `ALL CANDIDATE STRUCTURE AUDITS COMPLETE`, Python exit code 0, `SOURCE HASH UNCHANGED: YES`, and `NO WORKBOOK WAS MODIFIED`. The initial openpyxl Data Validation warning was non-fatal and did not prevent the completed audit.

### Acceptance and selection boundary

The 2026-11-12 authoritative handoff records:
- CC as the trusted synthetic source and read-only native probe A as frozen;
- synthetic native Save/reopen A as frozen;
- recurring AssignmentID migration A as frozen, with 179 immutable IDs independently verified.

These are evidence/reference artifacts, not writable targets. The recurring-ID workbook has the extra identity sheet and is the best reference for that bounded migration proof; the native-write workbook is the reference for its separate synthetic Save/reopen probe; the CC source/probe pair is byte-identical. None is approved as the full synthetic integration workbook. The whole-workbook synthetic fixture remains NOT CERTIFIED because all patient/staff-bearing surfaces must be proven synthetic and consistently keyed.

### New read-only VBA inventory tooling

Added:
- `Python/tools/audit_vba_project_inventory.py`
- `Tests/test_audit_vba_project_inventory.py`
- `Documentation/H10_VBA_PROJECT_INVENTORY_READONLY_RUNBOOK_20261115.md`

The new native Excel inspector disables macros/events, sets calculation to manual, opens XLSM read-only, lists VBA component names/types/code-line counts and UserForm control names/types, prints no code/cell values/captions, deduplicates byte-identical files and verifies source hashes before/after. It does not save or edit workbooks. VBA project access failures or locked projects are hard stops; do not change Trust Center/AccessVBOM settings to force access.

**Test status:** the last full source regression (1,813 passed, 1 skipped, 26 warnings, exit 0) predates the new VBA inventory tool. The new focused tests and full regression after these commits are NOT YET RUN. Do not report them green until operator evidence is received.

### Next operator gate

Follow `Documentation/H10_VBA_PROJECT_INVENTORY_READONLY_RUNBOOK_20261115.md`: pull latest branch, run the two focused auditor test files, then full source regression, then the read-only native VBA inventory for the four manifest paths. The duplicate CC file will be skipped automatically. If Trust Center access is unavailable, preserve the stop and upload the TXT; do not change security settings. No native write, Save, migration, candidate selection or whole-workbook certification is authorized.


### VBA inventory implementation hardening after initial documentation

The VBA inventory was further hardened before operator execution:
- `audit_one` now verifies source hash/size after every attempted native inspection, including controlled failures; closes without saving; and sets Excel calculation to manual before opening to avoid open-time formula recalculation.
- Duplicate handling now preserves the original audit failure status rather than treating a duplicate of a failed candidate as green.
- An unavailable UserForm control type marks the inventory incomplete (exit 3).
- The runbook's focused pytest invocation now captures stdout/stderr to its own TXT under a temporary `$ErrorActionPreference = "Continue"`, then checks the actual Python exit code.

Relevant source/test commits: `4b24f64c` (inventory tool), `19b730c6` (tests), `19df00b2` (failure/hash/calculation hardening), `07f79294` (duplicate failure preservation), `7a1c01d6` (duplicate-status test), `c0b3f7c0` (UserForm type fail-closed). Runbook updates: `0613de1e`, `6ff8e043`, `21319114`.

Current source/test status remains **NOT RUN after these commits**. The prior 1,813-pass full regression predates this tool and must not be reused as its acceptance evidence. Latest code commit for the VBA inventory: `c0b3f7c0`; all later commits in this sequence update documentation only.


## 2026-11-15 — Project-wide delivery register

The consolidated, evidence-based project delivery report is `Documentation/PROJECT_DELIVERY_MASTER_REPORT_20261115.md`. It brings together the H2–H10 delivery evidence, Recovery track, outpatient/registration/recurring-change workstreams, real-source admission, afternoon scope, protection gates and remaining release blockers. It does not replace specialist proofs or change any frozen artifact. Use `Documentation/CURRENT_HANDOFF.md` for the exact next operator action.


## 2026-11-15 — latest verified continuation and actual demo input boundary

The 20:50 expanded H10 call-edge gate is accepted: 45 focused tests passed; full regression 1,856 passed / 1 skipped / 26 warnings; three distinct H10 workbooks completed read-only VBA and structure audits with unchanged hashes and no modifications. See `CURRENT_HANDOFF.md` and `PROJECT_DELIVERY_MASTER_REPORT_20261115.md`.

The product target is a single functional Excel demo using the operator-reviewed patient/treatment source. The two private source and identity-review workbooks are already available in the conversation context. Do not ask for re-upload and do not commit either workbook or patient data. The operator's A:T mapping remains the core schema; U/V contain additional afternoon and pool fields. Identity suggestions are not approvals. Build and test the source-shaped parser/projection with synthetic data first; real-patient admission remains separately gated.


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


## 2026-11-16 — F/U/V read-only gate, 1,991 tests GREEN

**Concrete new gate, not a synthetic replacement for the user's real-data DEMO:** `Python/rehab_excel/trial_special_fields_audit.py`, `Python/tools/audit_real_manual_specials.py`, `Tests/test_trial_special_fields_audit.py` provide the first **single-source-only** private read-only inspection of actual manual A:V columns **F** (reclining, known 30-minute ΦΘ-related operation), **U** (legacy afternoon, no auto-import) and **V** (morning pool). No review/clinical approval/clean XLSM required for this *diagnostic* first command. It checks original source sheet/header, SHA-256 before/after, cached and raw Excel formula values (never omit formula-only cells), missing-patient source rows, time parseability, all counts and no PII emitted. The tool **never creates sessions, guesses F/V weekdays or imports old U**. It returns 0 if F/U/V truly empty (not whole source approval), 3 for pending real special records, 2 on safety stop. Exact local command, interpretations and privacy gates: `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md`.

**Verified code evidence:** `95157af7dbb32c6cb8aca210c503ffbb7f63be17`, GitHub Actions `43804737124`, **1,991 passing pytest in both Windows and Linux**, all four full/synthetic CI jobs successful. This is not proof that the real private manual has been scanned or a working native Excel was built. **Real-data DEMO remains NOT DELIVERED.** F/V day/provider/duration-aware scheduling approval, U source scope and independently reviewed outpatient programme remain blockers. Next meaningful step is run the read-only audit against the operator-approved private CURRENT manual snapshot, capture only totals, finish explicit F/V mapping and standalone outpatient source, then eligible private native seed and full side-by-side weekday/operator acceptance. Frozen H2–H10 proof artifacts remain untouched.

 
## Latest continuation checkpoint — 2026-11-25

Before continuing DAILY_INPUT verifier work, read:
- `DIOMIL_MASTER_CONTINUITY_GUIDE.md` §4.226–§4.227
- `CURRENT_HANDOFF.md` (authoritative latest handoff section)
- `DAILY_INPUT_INTERNAL_VERIFIER_REVIEW_20261125.md` §12
- `DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md`

Latest code/test HEAD: `e7ce7526de2c000cff9410819267f73fdc6ffc79`; later commits may be documentation-only. The latest full cross-platform CI is run `45657152050`, 2848 passed on Windows and Ubuntu, with 35 and 34 warnings respectively. The local focused Windows suite is 214 passed, 1 symlink-specific skip, no warnings summary.

The newly added native MASTER diagnostic is partial and fake-COM tested only. It does not complete PROJECTIONS_REOPEN. Continue with independent native therapist AssignmentID extraction and full projection/source-ledger linkage. Gate 2 remains closed.

 
The next therapist verifier must reuse the stable recurring ID and daily-state source chain mapped in `DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md` §7. The visible THERAPIST_DAILY sheet has no AssignmentID field; do not infer stable identities from its labels or occupied-clock totals.

 
The native THERAPIST_DAILY comparison primitive is now present. Before adding a phase adapter, implement the complete 126-cell expected-grid builder from the existing unified schedule, stable AssignmentID, daily-state/replacement and patient-centric renderer sources. The primitive and its exact limitations are in `DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md` §8. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED.

 
The expected THERAPIST_DAILY grid builder is now available and cross-platform synthetic CI is green (run `45660810319`, 2866 passed on both Windows and Ubuntu). It requires caller-supplied independently validated stable-ID schedule/session/state facts and does not itself verify DAILY_INPUT/ledger/replacement sources. Next is a composite native diagnostic that binds this expected grid to the full-grid COM readback and exact operation/source/date evidence. See design doc §9.
