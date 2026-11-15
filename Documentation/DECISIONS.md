# Project Decisions

This file records business and architecture decisions that should survive loss of chat/session context.

## Patient identity

### Decision
Use two separate identifiers:

- `PatientID`: permanent internal system identifier, generated automatically.
- `HospitalMRN` / `ΑΜ Νοσοκομείου`: optional hospital registry number, entered when known and editable later.

### Rationale
`PatientID` is used as the stable key across schedules, absences, replacements, history and future related records. It must not be replaced later by another identifier. Hospital MRN is external business data and may be unknown at first.

## Patient type

### Values
- `Εσωτερικός`
- `Εξωτερικός`

### UI rule
The patient registration combo box defaults to `Εσωτερικός` to minimize routine clicks. The user can switch to `Εξωτερικός` from the drop-down.

### Operational ratio
Outpatients are expected to be about 15% max of the total patient population. This is a practical operating ratio, not a hard validation ceiling.

## Outpatient scheduling semantics

### Core rule
An outpatient is fully real for scheduling and accounting. Patient type affects visibility and presentation, not operational counting.

### Must count normally in
- therapist daily workload
- provider capacity
- time-share / productivity accounting
- recurring treatment schedules
- replacement calculations
- daily therapist schedules

### Visibility
Outpatients should not appear in inpatient-only patient views or the inpatient patient planner.

Outpatients should appear in `THERAPIST DAILY`.

### Recurring schedule storage
`PATIENT_PLANNER` remains the inpatient/hospitalized-patient planner and must not become the visible home of outpatient schedules.

However, outpatient recurring appointments still need an authoritative persistent source. They must not exist only as text in `THERAPIST DAILY`, because workload, capacity, replacements, cancellations and future-day materialization all depend on the recurring source data.

Architecture decision:
- keep inpatient recurring rows in `PATIENT_PLANNER`
- store outpatient recurring rows in a separate authoritative source, provisionally named `OUTPATIENT_SCHEDULE`
- the scheduling engine reads/merges both sources into the same `BaseScheduleEntry` stream
- after merging, inpatient and outpatient sessions use the same operational scheduling engine
- `THERAPIST DAILY` remains the only normal operational view where outpatient names are shown

`OUTPATIENT_SCHEDULE` is an authoritative storage sheet, not an end-user workspace. The normal application must keep it hidden/protected and must not require the user to navigate to it, know Excel row numbers, or edit its cells directly.

Normal create/edit access for outpatient recurring schedules must come from the application's top-level UI. The intended UX is a management view/form that lists outpatient schedule entries in business terms such as patient, treatment, time, days and therapist, with actions such as create and edit. Internal identifiers such as `outpatient:<row>` may remain implementation keys but must not be required input from the user.

Direct access to `OUTPATIENT_SCHEDULE` may remain available only for controlled development/support diagnostics, not as part of the ordinary clinical workflow.

### Application-style UX do's and don'ts

The workbook is the execution/storage container, but the user-facing product should behave like dedicated software rather than a collection of Excel sheets.

Do:
- expose normal work through the top application menu, buttons, forms and guided workflows,
- show business concepts such as patient, treatment, therapist, date, time and status,
- let the system carry internal row numbers, sheet names, IDs and storage details invisibly,
- provide next-step guidance after successful actions when that helps the workflow,
- keep technical/data sheets hidden or protected unless an explicit support/development action requires them,
- make create, edit, replacement, conflict resolution and review actions reachable without showing the Excel grid or native sheet navigation.

Don't:
- require the user to know or type worksheet row numbers,
- require the user to know internal sheet names such as `OUTPATIENT_SCHEDULE`,
- expose `outpatient:<row>`, internal English command strings, backend lifecycle details or implementation mechanics as normal user-facing information,
- send the user to hidden data sheets merely to select or edit a record,
- design a workflow that assumes the Excel ribbon, worksheet tabs or raw grid are visible,
- use success messages to report technical implementation facts when an operational instruction or next action is more useful.

Developer-only smoke/debug interfaces may temporarily expose internal details, but they must be clearly treated as test scaffolding and must not become the final user workflow.

### Strict visual / clinical rule
An outpatient is **light blue only** in `THERAPIST DAILY`.

An outpatient:
- is never infectious/yellow
- never receives robotic treatment / robotic extra programme
- must not inherit yellow, pink, or any other patient colour from another patient who shares the same recurring therapist/time slot on other weekdays

When a therapist/time slot is shared by different patients on different weekdays, `THERAPIST DAILY` presentation is determined by the patient actually active on the concrete date.

If the same effective daily cell would contain both an inpatient and an outpatient simultaneously, cell-level colouring is ambiguous and the renderer must reject the situation rather than guess a colour.

### Daily postponement / no-show rule
The recurring outpatient appointment remains in the base schedule, but any specific occurrence may be cancelled for that date without deleting or moving the recurring slot.

A daily cancelled occurrence must:
- be marked with a reason, distinguishing at least `department_cancelled` from `patient_no_show`
- stop counting as delivered treatment for that date
- release the therapist timeslot so it can be used for replacement work if needed
- preserve the original recurring appointment for future dates
- remain auditable in daily history/statistics

This is a daily operational override, not a change to the patient's base programme.

## Historical statistics and finalized-day corrections

### Finalization
Historical daily workload becomes final only through an explicit application action `Κλείσιμο ημέρας`. The action must be blocked while unresolved/orphan sessions remain or while any concrete-date session lacks a final operational outcome.

### Persistence model
Historical storage is hybrid:

- finalized daily snapshots support fast weekly/monthly reporting,
- append-only events preserve the audit trail and correction history.

The primary workload metric is the actual final operational workload. Planned workload may be retained only as secondary comparison context.

### Correction rule
A finalized day must not be silently reopened and destructively rewritten.

Use a dedicated user action such as `Διόρθωση κλεισμένης ημέρας`.

A correction must create an auditable amendment event, pass the normal validation/finalization gate again, create a new finalized snapshot version, preserve the previous version as superseded, and link the new version to the previous one. Reporting uses the latest current-final snapshot; audit views may display the full version chain.

Failed corrections must not partially replace the current-final historical state.

## Global replacement optimization

### Decision
When multiple therapist absences create displaced sessions on the same day, do not solve replacements greedily one therapist or one patient at a time. Record all known therapist absences first, collect all displaced sessions, and evaluate the full set of patients and candidate placements together before committing the proposed assignment.

### Hard feasibility constraints
A candidate placement is invalid when any of the following applies:

- the patient has another treatment occupying that time,
- the patient has an explicitly recorded absence or unavailable hour,
- the therapist is absent, blocked, excluded, not qualified for the treatment, or otherwise unavailable,
- the placement exceeds an applicable provider/student capacity rule without an explicitly approved exception,
- the treatment requires a service or equipment area that is not operating at that time,
- an existing daily fact or other locked operational rule makes the placement invalid.

Pool and robotic services have their own operating hours. Those hours must be stored as persistent configuration and remain in force until deliberately changed; the solver must never assume that these services operate at every ordinary treatment time. Robotic appointments may move only within the configured robotic operating hours and must respect known therapist/provider assignments. More detailed monthly robotic restrictions can be added later without blocking the initial global solver.

Do not introduce mandatory, detailed patient-availability data entry for every patient. Initially derive constraints from the existing schedule, treatment conflicts, explicit patient absences, provider policies, and the persistent operating-hour configuration.

### Global objective and preferences
The solver must consider the complete candidate set for all displaced patients on the concrete date. Its intended priorities are:

1. Never violate a hard feasibility constraint.
2. Maximize the number of displaced patients placed and minimize unresolved orphans.
3. Protect scarce placements globally. A flexible patient must not consume a provider/time slot that is the only feasible option for another patient when a better complete assignment exists.
4. Prefer keeping each patient at the original time, unless doing so harms another patient's unique option or conflicts with the important adjacency of a reclining treatment.
5. Preserve the relationship between physiotherapy (ΦΘ) and a reclining treatment where possible. The reclining treatment should preferably remain in an adjacent or nearby hour, and unnecessary waiting between a patient's treatments should be reduced.
6. Keep provider workload as balanced as practical. A difference of about one treatment-hour may be unavoidable; a two-hour imbalance is considered significant and should be avoided when a feasible alternative exists.
7. Minimize unnecessary changes to the patient's time and therapist, while considering the preceding constraints and preferences.

An explicit example of the scarcity rule is locked: if patient A can use 09:00, 10:00, or 11:00 while patient B can use only 09:00, protect 09:00 for B and assign A elsewhere whenever a complete feasible assignment exists.

The 13:00 timeslot is an additional **weak soft preference to avoid**, not a forbidden time and not a hard constraint. Among otherwise acceptable assignments, prefer solutions that do not use 13:00. This preference is lower priority than placing all patients, protecting unique feasible options, preserving the original time where appropriate, keeping physiotherapy/reclining treatments near each other, and avoiding significant workload imbalance. If 13:00 is needed to produce an otherwise valid complete assignment, use it rather than leave a patient unresolved or violate a higher-priority rule.

The precise relative weights between same-time preservation, reclining adjacency, workload balance, and the weak 13:00 avoidance preference must be validated with real scheduling examples before they are frozen as numeric scores. The business direction is already clear: protect unique options first, preserve the original time where feasible, treat reclining adjacency as important, avoid large workload imbalance, and avoid 13:00 only when doing so does not compromise those higher priorities.

### Batch workflow
- Enter all known therapist absences for the day before resolving displaced sessions.
- Generate candidate placements for every displaced patient.
- Solve the combined assignment as one global problem; do not commit early local choices that can make the remaining set impossible.
- Present the resulting proposal and any unresolved patients for review before finalizing the daily state.

A weighted matching, min-cost flow, or constraint solver is acceptable if it is deterministic, testable, and follows the agreed priorities.

### Double-booking / over-capacity exception
Double-booking or an over-capacity placement is not a normal automatic solution. Consider it only when no acceptable regular assignment remains. In that case, the system must ask the user explicitly which remaining patient could share a treatment time with another patient and with which therapist, show the relevant warning, and require confirmation for each exception. Never silently create such an assignment or pre-approve it in the optimizer.

### Selective clinical and operational knowledge
Some important knowledge is case-specific and may include therapist/patient physical suitability, working relationship or compatibility, and how well an older or medically burdened patient tolerates waiting. Do not require a large set of new fields for every patient. Allow such exceptions to be recorded selectively when useful, or handled through explicit user review. The solver must not invent or infer these judgments when they have not been recorded.

### Manual override
Manual changes remain available. Automatic optimization should produce the best feasible global proposal while preserving the user's authority over exceptional clinical and operational decisions.

## Registration UI

- Central UserForm: `frmRegistrationMenu`
- Patient UserForm: `frmNewPatient`
- Current patient form is preview-only and must not write to the workbook yet.
- Room options come from `SETTINGS` column H.
- Patient status options come from `SETTINGS` column F.
- Inpatient-only controls are disabled/cleared when type changes to outpatient.

## Daily operational UI terminology

User-visible daily operational fields must use the agreed Greek business terminology only. Internal technical identifiers, source labels, command names or English implementation strings must never leak into normal Excel UI fields such as `Αιτία`.

For replacement rows created by the `DAILY_INPUT` therapist-absence workflow, legacy technical values such as `DAILY_INPUT therapist absence` or `temporary replacement` are rendered to the user as `Απουσία θεραπευτή`. Internal immutable IDs may remain technical because they are not user-facing labels.

## Excel/VBA compatibility

The target Excel/pywin32 installation rejected programmatic COM writes to some MSForms `Width`/`Height` properties.

Decision:
- use COM only to create the form and controls
- perform final sizing/styling from VBA `UserForm_Initialize`

This pattern is now the preferred approach for future dynamically installed VBA UserForms unless a later compatibility test proves a better method.

## Development safety

- Never modify the baseline `.xlsm` directly during development.
- Create preview copies first.
- Verify source hash remains unchanged.
- Preserve VBA project content.
- Keep entity-specific backend validation authoritative. UI validation is convenience/presentation only.

## Documentation maintenance rule

Documentation is part of the implementation, not a cleanup task after implementation.

Whenever any meaningful development occurs, update the appropriate project records in the same work cycle. This includes:

- implemented behaviour or changed semantics,
- new tests or regression findings,
- compatibility workarounds,
- current validation state and native smoke results,
- defects, risks and technical debt,
- newly discovered necessities,
- ideas or requirements that we decide should later be built,
- functionality that should be removed, retired or replaced,
- fixes that are known but intentionally deferred,
- changes to implementation order or project priorities.

Use the files by purpose:

- `Documentation/DECISIONS.md` for durable business/architecture rules and agreed constraints,
- `Documentation/PROJECT_STATUS.md` for current progress, open risks, validation state and immediate roadmap,
- `Documentation/CURRENT_HANDOFF.md` for the exact resume point, latest checkpoints and commands/scenarios needed to continue safely,
- `Documentation/IMPLEMENTATION_MAP.md` when ownership/location of implemented capability changes,
- `CHANGELOG.md` for user-visible or implementation-relevant changes that belong in release history.

Do not rely on chat history as the only record of a decision, requirement, idea, removal, deferred fix or newly discovered need. If it may affect future implementation, testing or recovery after context loss, it must be written into the repository documentation.

## Durable decisions — 2026-11-12: permanent therapy identity and history

- A recurring treatment must have a durable, opaque **AssignmentID**, not `weekly:row`, `outpatient:row` or `planner:row:treatment`, before events can be committed. Existing transient IDs are allowed only in read-only legacy previews.
- Planner IDs are keyed by authoritative `(PatientID, Treatment)` in technical `RECURRING_ASSIGNMENT_IDS`; granular weekly and outpatient rows carry explicit `AssignmentID` column G. Reject duplicates, blank identities, mismatched source rows and unapproved/ambiguous identity mapping. Never guess identities from names, clock time or row index.
- A permanent therapy change is **effective-dated and append-only**; events never rewrite original recurring base or already-finalized historical service-day records. Event IDs must be unique; chronological and same-day conflicts fail closed. Real reversal must be a tested compensation/versioning workflow, not deleting history.
- Provider capacity is **actual approved distinct timeslots per therapist**, never a universal default 6 or raw count of patient lines. Grouped patients may share one therapist slot only with separately audited approval. Patient collisions across specialties/weekdays must fail closed.
- Production clinical workbook and immutable baseline are never modified by development smokes. Source CC and accepted A native candidates are frozen. Native installation requires exact-source preflight, exclusive fresh output, no overwrite, disabled macros for structural tools, a single controlled Save, independent reopen, source SHA verification and no-retry-on-failure.
- 2026-11-12 operator-confirmed native proof of 179 IDs is **not** proof of event writes, daily Apply/Restore, clinical workflow correctness or real roster admission. As of that checkpoint `recurring_change_event_source.py` is read-only; no event append or rollback is approved.
- Patient-identifying trial staging and the `ΠΡΟΓΡΑΜΜΑ` source must remain private and off public GitHub/CI. The real-data admission gate blocks unknown IDs, unresolved source rows and unadjudicated patient/provider collisions. Never autoapprove or silently discard held appointments.
- Management summary UX must be visually accepted, capacity mathematics audited and exports whitelisted/macro-free/privacy-tested before external sharing. Internal XLSM is never a management share artifact.
- Documentation is a release safety gate: update **ONBOARDING, CURRENT_HANDOFF, PROJECT_STATUS, IMPLEMENTATION_MAP, DECISIONS, CHANGELOG, and subsystem/native runbooks** alongside meaningful source work. Include full exact copyable PowerShell commands in the relevant operator runbook, tested/not-tested scope, safety-stop contingencies, commits, frozen artifacts and next steps. The latest canonical handoff is `Documentation/AUTHORITATIVE_RESUME_20261112.md`. Historical "next" sections may remain but must be explicitly superseded rather than executed.



## 2026-11-12 additional recurring-event transaction guard

A permanent change may be *proposed* only after source SHA256, immutable AssignmentID, exact operator-seen assignment fingerprint, unique new EventID, valid prior ledger and later-than-prior per-assignment effective date checks all pass. Reuse the actual current event history in shared collision/capacity preflight. The proposal is **not a write authorization**, must be revalidated against an unchanged source immediately before native copy/write, and may never mutate frozen synthetic A candidates. Fail closed if existing event rows are malformed or event IDs duplicate. Any future native writer needs separate atomic rollback and independent persisted read-back before UI enablement. See `recurring_change_append_plan.py` and `Documentation/AUTHORITATIVE_RESUME_20261112.md`.


### 2026-11-13 real-source demo preflight preparation (documentation-only)
Read `Documentation/REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md` for the latest private-input preflight, safe single-block PowerShell, stop rules, precise unsupported import boundary and weekend native acceptance plan. Tracked `trial_admission_gate.py`, `preflight_trial_excel_sources.py` and `audit_pilot_source_readiness.py` were inspected, but **no code, live private source, Windows Excel, or full tests were executed in this checkpoint**. No real-data XLSM was produced. The read-only gate is NOT an importer; strict gate still blocks on unresolved holds, overlaps and shared slots. Next work requires confirmed local private source revision and separate eligible pilot intake, then actual supported native importer engineering and one-shot proof. Do not reuse synthetic CC or frozen artifacts; keep clinical values off public GitHub.


### 2026-11-13 source-only trial import planning checkpoint
Added `Python/rehab_excel/trial_import_plan.py` (commit `b0e9d48`) and `Tests/test_trial_import_plan.py` (commit `0a9bc4c`). The plan reuses strict read-only combined preflight and exact approved `SourceWeeklyEntry` admission, cross-checks counts, binds both source SHA256 values, rejects changed sources, and hides private entries from its object repr. **It does not write an XLSM, implement a native importer, or authorize clinical use.** The new test file has six synthetic cases, but **tests/CI have NOT yet been observed passing**. Before native integration run `python -m pytest Tests/test_trial_import_plan.py Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q` in one PowerShell fail-fast block; stop on red. No private patient data uploaded. Current real staging remains unverified. More work is needed to build the new isolated native importer and independent readback.


### 2026-11-13 verified-schema weekly projection SOURCE checkpoint
Source-only modules added: `Python/rehab_excel/trial_weekly_projection.py` (`f90d4b6`) and seven synthetic tests `Tests/test_trial_weekly_projection.py` (`f2e6268`). This pure bridge maps independently approved `TrialImportPlan` entries into existing `PROGRAM_WEEKLY_SLOTS` six-column headers, with strict inpatient ID membership, no silent omissions, treatment/provider checks and patient-time collision rejection. The synthetic roundtrip test invokes the **existing `read_weekly_slots_sheet`** against an in-memory workbook. No real patient data, Excel file output, COM write, native acceptance or approved clinical import yet. New tests are **not confirmed green** unless a subsequent CI/native result explicitly proves that. Run `python -m pytest Tests/test_trial_import_plan.py Tests/test_trial_weekly_projection.py Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q` as one fail-fast PowerShell gate. Still missing: approved patient registry detail mapping, a real source-bound native writer, independent readback, privacy and operator proofs. Do not treat `PROGRAM_WEEKLY_SLOTS` alone as a complete operational demo.


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


## 2026-11-15 — integrated demo source admission decisions

- The target demo must be built from the operator's reviewed patient/treatment source and agreed mapping; a synthetic-only workbook is a test fixture, not the final demo target.
- The source schema includes the supplied A:T mapping and additional U/V afternoon/pool fields. Do not silently omit U/V.
- Identity matching suggestions, including exact-name suggestions, are not identity approval. No patient row may be activated/imported without explicit confirmation and conflict/source-hold resolution.
- Private patient source/review workbooks and patient-level audit output must remain outside the public repository and public logs.
- The engineering order is synthetic A:V fixture -> parser/projection tests -> isolated native write/readback design -> private admission review -> operator acceptance. Frozen H10 evidence workbooks remain immutable.


## 2026-11-15 — A:V real-source demo schema design

Added `Documentation/TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md` as the current source-to-demo design contract. It records the operator's A:T mapping, the additional U/V afternoon/pool fields, the boundaries of the existing read-only admission/import-plan/weekly-projection modules, and the required architecture for source alignment, patient registry projection, full recurring schedule projection, synthetic end-to-end tests and later native Excel acceptance.

Important unresolved rules are explicit rather than guessed: blank weekday semantics; missing A context handling; U/V operational scope; and fresh operator instructions for reclining replacement behavior at implementation time. The private identity review remains unapproved; no real-patient import or native integrated demo is authorized. Keep source workbooks and all patient-level facts outside GitHub.
