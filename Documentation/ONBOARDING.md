# Project Onboarding

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
