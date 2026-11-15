# Project Decisions

## ΝΕΟ ΚΡΙΤΗΡΙΟ ΑΠΟΔΟΧΗΣ 2026-11-16: REAL-DATA DEMO

**Οριστική νέα απόφαση χρήστη:** Ένα DEMO με τυχαίους ασθενείς/ώρες/θεραπείες **δεν επιτρέπει αξιολόγηση**. Το παραδοτέο επιχειρησιακής πρακτικής δοκιμής πρέπει να διαβάζει **τα πραγματικά δεδομένα του τρέχοντος χειροκίνητου προγράμματος**, ακριβώς ανά ασθενή/θεραπεία/ημέρα/ώρα/θεραπευτή, σε ιδιωτικό ξεχωριστό XLSM. ΔΕ–ΠΑ μόνο, μηδέν θεραπείες Σ/Κ. Ο αρχικός χειροκίνητος κορμός παραμένει αμετάβλητος και authoritative. Synthetic fixtures επιτρέπονται για μηχανικά tests, όχι ως demo-to-user. Πηγαία blockers/F/U/V **δεν** παραλείπονται σιωπηλά. Οι κλινικές/identity εγκρίσεις πρέπει να έχουν δοθεί από εξουσιοδοτημένο πρόσωπο σε ιδιωτική υποδομή. Το ολοκληρωμένο demo σημαίνει native COM + πλήρη 1:1 αντιπαραβολή και επιβεβαιωμένες user operations, όχι pytest GREEN. Κατά την παράδοση ενημερώνονται όλες οι **ζωντανές** αναφορές παραδοτέων/hand-off/υπολοίπων χωρίς επανεγγραφή frozen historical proofs. Αναλυτικός κανόνας: `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`.

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

 
## 2026-11-15 — source programme operating weekdays

- There are no treatment sessions on Saturday or Sunday.
- The operating week is Monday–Friday (`ΔΕ, ΤΡ, ΤΕ, ΠΕ, ΠΑ`).
- For source import only, a populated treatment time/value with a blank corresponding weekday cell means recurring Monday–Friday.
- A blank treatment time/value creates no session, regardless of the weekday cell.
- Explicitly stated weekday patterns remain authoritative. Never generate weekend sessions.
- Implement via the source importer’s explicit `implicit_daily=True` path; do not change the general parser default globally.


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
