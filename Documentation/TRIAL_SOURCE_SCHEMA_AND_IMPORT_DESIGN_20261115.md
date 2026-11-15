# DIOMIL — Real-source demo schema and import design

## ΕΠΙΧΕΙΡΗΣΙΑΚΗ ΕΝΗΜΕΡΩΣΗ 2026-11-16: ΠΡΑΓΜΑΤΙΚΟ DEMO

**Project user confirmation, 2026-11-16:** No synthetic patients/times/treatments in the evaluation deliverable. Only a **fresh current private hand-kept A:V source**, approved identity review, strict complete source match and approved clinical fields can seed the separately isolated XLSM. F/U/V require full semantics and block the source when present. Exact ΔΕ–ΠΑ and weekdays with no clock must obey original contract; zero weekend therapy. Preserve source-row and provenance; do not invent room/doctor from A, infer MRN/illness from formatting, reuse IDs or silently omit therapy. All private workbook bytes stay off GitHub. Native source/schedule readback and private manual comparison are still PENDING. See `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md` for the updated release acceptance plan.

Date: 2026-11-15  
Branch: `work/afternoon-medical-scope-20261115`  
Status: source schema inspected locally; design only; no patient import or workbook write authorized.

## 1. Product target

The target is one functional, integrated Excel demo based on the operator-reviewed patient/treatment programme, using existing proven Diomil operations. A synthetic workbook is a test fixture and engineering gate, not a substitute for the requested real-source demo.

All real source/review workbooks remain on institution-controlled storage. Never commit workbook bytes, patient names, IDs, room/doctor values, patient-level logs, or private paths to GitHub or public CI.

## 2. Source programme field map

The operator supplied the A:T mapping, and the inspected programme sheet contains additional U/V columns. Preserve the complete source width; do not silently discard U/V.

| Column | Meaning | Import interpretation |
|---|---|---|
| A | Clinic / ward / doctor context | Preserve source context and provenance; blank cells are not automatically forward-filled. |
| B | Patient name | Display identity only; never use name alone as PatientID. |
| C | ΦΘ time | Recurring physiotherapy start time. |
| D | ΦΘ days | Explicit weekday pattern for ΦΘ. Blank-day semantics must be confirmed before parser admission. |
| E | ΦΘ therapist | Provider associated with ΦΘ. |
| F | Reclining time only | A ΦΘ-related operation at a separate time, 30 minutes. It is not a separate ordinary specialty. Its recurrence/day association and replacement behavior must be confirmed before writeback. |
| G | Ε/Θ time | Occupational therapy start time. |
| H | Ε/Θ days | Occupational therapy weekdays. |
| I | Ε/Θ therapist | Occupational therapist. |
| J | Λ/Θ time | Speech therapy start time. |
| K | Λ/Θ days | Speech therapy weekdays. |
| L | Λ/Θ therapist | Speech therapist. |
| M | Ψ/Λ time | Psychology start time. |
| N | Ψ/Λ days | Psychology weekdays. |
| O | Ψ/Λ therapist | Psychologist. |
| P | ΕΦΑ hours | Special physical education time(s). |
| Q | ΕΦΑ days | Special physical education weekdays. |
| R | Robotics time | Robotic treatment start time(s). |
| S | Robotics days | Robotic treatment weekdays. |
| T | Robotics therapist | Robotic therapist. |
| U | Afternoon | Additional afternoon source field; not yet mapped to an accepted operational model. |
| V | Pool | Additional pool source field; not yet mapped to an accepted operational model. |

The source includes Excel numeric and text time representations, grouped weekday patterns, slash-separated values and incomplete/notes-only cells. Reuse the tested `rehab_core.programme_mapping` semantics where they match the source. Any source shape outside that contract must become a review hold, not a guessed value.

## 3. What the existing trial pipeline does

- `trial_identity_review.py` creates a separate identity-suggestion copy. Suggested old IDs are explicitly unconfirmed and do not populate the authoritative PatientID/confirmation columns.
- `trial_admission_gate.py` consumes the normalized review workbook. It requires explicit PatientID confirmation, rejects duplicate IDs, unresolved source holds, cross-specialty overlaps, invalid weekly entries and unapproved shared therapist/day/time membership.
- `preflight_trial_excel_sources.py` is read-only and compares a review XLSX with a separate eligible XLSM intake candidate. It does not import.
- `trial_import_plan.py` is a read-only plan over approved weekly entries and binds hashes of the review and intake workbooks. It does not yet bind the original A:V source workbook.
- `trial_weekly_projection.py` projects only the weekly-slot sheet. It does not project the complete patient registry, demographics/context, all source columns, afternoon/pool fields, or a full operational MASTER programme.
- `weekly_slots_source.py` reads the optional six-column `PROGRAM_WEEKLY_SLOTS` schema and overlays granular entries over legacy patient/treatment pairs. It is not a complete source importer.
- `BaseScheduleEntry` currently has no explicit duration field. The 30-minute reclining operation must not be represented as an independent durationless generic slot if downstream interval/capacity logic needs its duration.

Therefore, the existing source-only tools are necessary admission components, not a complete importer or whole-workbook integration.

## 4. Required target architecture

1. **Private source alignment:** read the original A:V programme and normalized review workbook read-only; verify source row/name/context provenance and every source-cell reference; bind both source hashes. Never print identities.
2. **Identity admission:** explicit operator-confirmed PatientID for every admitted patient; suggested IDs remain suggestions until confirmed. Do not auto-match by name.
3. **Clinical/source review:** unresolved notes, malformed times/days, patient overlaps and shared-provider slots remain held until explicitly adjudicated. The current strict gate requires zero unresolved holds/overlaps; no partial admission is silently allowed.
4. **Canonical patient projection:** create complete Patient registry records, including clinic/ward/doctor context and inpatient/outpatient classification only where supported by explicit source facts.
5. **Canonical recurring schedule projection:** preserve each treatment/day/time/provider as an individual source-bound entry. Do not collapse different weekdays or multiple clock groups into one lossy patient cell.
6. **Additional treatment streams:** model reclining as the specified 30-minute ΦΘ-related operation; map afternoon and pool only after their operational rules and source semantics are confirmed.
7. **Synthetic end-to-end fixture:** use synthetic patient/provider values to exercise A:V parsing, identity gates, holds, weekly projection, patient registry projection, shared slots, collisions and source-hash checks.
8. **Native integration:** only after source/tests are green, build a fresh unique XLSM from a separately eligible trusted source. Use native Excel for save/close/reopen and independent readback. Never modify frozen H10 proof artifacts.
9. **Acceptance:** verify patient identity, every treatment/day/time/provider, source traceability, no phantom sessions, history preservation, navigation and daily workflows. Record exact tested scope and limitations.

## 5. Current private-input observations (non-identifying)

The private review workbook is a proposal/review artifact, not an approval ledger. The patient identity confirmation field is still pending; exact-name suggestions remain unconfirmed; source notes and overlap holds remain unresolved. The review patient rows were cross-checked against the source sheet by source row and patient name, and the two non-patient section/header rows are not patient records. This limited provenance check does not validate every proposed time/day/provider interpretation.

The source context column A is blank on some patient rows, and the review workbook preserves those blanks. Do not forward-fill ward/doctor context without an explicit rule. The source's two clinic sections are represented separately in the review workbook; do not infer clinic from the ward/doctor text.

## 6. Explicit unresolved operator rules

- What does a blank weekday cell mean for each treatment: every weekday, no scheduled days, or an incomplete source hold? The current review proposals expand some blank day fields to weekdays, but this is not accepted as a business rule until confirmed.
- Are U (afternoon) and V (pool) in scope for the first integrated demo, and what are their exact time/day/provider semantics?
- At the implementation point for reclining replacements, request the operator's fresh full instructions. Current known boundary: reclining is ΦΘ-related, separate time, 30 minutes; replacement time is evaluated as an interval and proximity to ΦΘ is preferred when feasible. Do not implement the optimizer from this abbreviated note alone.
- Confirm how missing A context should be handled; never silently inherit a prior ward/doctor value.

## 7. Safety and evidence boundary

No real patient row has been admitted by this design document. No real-data importer, writable destination, native Excel Save/reopen proof or integrated real-data demo has been produced. A positive source-only preflight is not a clinical approval or native acceptance. Keep source/review files, identity decisions, hashes and patient-level evidence private.


## 2026-11-15 — source alignment tooling added and CI accepted

The read-only provenance audit is now implemented in `trial_source_alignment.py`, with private local CLI `audit_trial_source_alignment.py` and synthetic tests. It compares source row/name/context and weekly source-cell references without printing identities or source paths. GitHub Actions run `43751221120` passed 200 focused tests on both operating systems and 1,869 full regression tests on both operating systems. This is source/test evidence only, not native Excel or import acceptance.

The uploaded private source/review pair was separately cross-checked without exposing patient values: source-row/name/context and staged source-cell references aligned; blank A context remains blank. Identity approval and all source/conflict holds remain unresolved. The next implementation should bind the original programme hash into the import plan and add complete patient-registry projection before any native writer.

 
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
