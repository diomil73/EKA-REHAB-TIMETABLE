# Implementation Map

## Implementation owner: unified private real-source DEMO admission (2026-11-16)

- `Python/rehab_excel/trial_private_outpatient_audit.py`: existing-schema read-only outpatient eligibility checks; does NOT import.
- `Python/rehab_excel/trial_private_outpatient_projection.py`: independently audited private outpatient registry and appointment projection with source SHA, exact weekdays, and no unapproved extra columns.
- `Python/rehab_excel/trial_demo_source_projection.py`: source-to-source PatientID collision gate and combined patient registry while preserving separate inpatient weekly and outpatient recurring schedule ownership.
- `Python/tools/prepare_real_data_practice_demo.py`: read-only dual-source `--outpatients` gate; native `--build` requires that independent source.
- `Python/rehab_excel/trial_demo_native_seed.py`: existing schema OUTPATIENT_SCHEDULE native COM write, combined weekly view, raw and production unified-schedule readbacks, original hash checks.
- `Tests/test_trial_private_outpatient_projection.py`, `Tests/test_trial_private_outpatient_audit.py`, `Tests/test_prepare_real_data_practice_demo.py`, `Tests/test_trial_demo_native_seed.py`: fictional regression, path/identity guards, source-only vs native distinction.

CI gate `70e1b88f6e3289bbc79fe411d4724e327622ff7f`, run 43809724956: 2,029 passing tests each on Windows/Linux, 4/4 jobs. **No actual private real-source native XLSM or operator acceptance.** Preserve daily MASTER VBA renderer; do not repurpose outpatient records into inpatient PATIENT_PLANNER. F/U/V mapping and active-MASTER discharge refresh remain independent open gates.


## 2026-11-16 — Executable private real-source DEMO preflight

**New executable interface:** `Python/tools/prepare_real_data_practice_demo.py` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Default operation is **read-only full real-source validation** for an authorized current A:V manual XLSX, confirmed review XLSX, entirely empty application-ready VBA XLSM and independently approved private `clinical_approvals.json`. Uses the strict production source alignment, all-patient approved ID check, original weekday/time/provider reconciliation, unresolved F/U/V stop gates, existing native shell checks and SHA-256 stability of the clinical-approval JSON. Every input must remain off Git checkout. No real patient name, ID, room, doctor, MRN, source path or clinical exception content is emitted to stdout/stderr on handled errors. `--build --output <PRIVATE_*_DEMO_*.xlsm>` is an *explicit* Windows + native Excel/COM, exclusive-new-file second step and repeats the same approval gates. **No real source file has been run; no real-data XLSM exists, and end-to-end practical acceptance is still pending.** Previous synthetic/H8 workbook results do not count as the operator's comparison. First actual action is an authorized private preflight with current manual and approved clinical review, not a random synthetic seed. Native failure or changed approvals after a native operation requires freezing the candidate, not treating it as accepted. See the dedicated operator CLI runbook for exact commands/JSON schema.

**Current blocker:** F/U/V and independently sourced outpatient bookings remain unavailable to the current *inpatient-only* A:V admission bridge. If occupied, the complete real-data source fails closed, never silently drops treatments. Do not alter frozen H2-H10 evidence. Update all living handoff/status/test and delivery documents again with actual source-specific counts and results **only after** the first private native execution, without publishing PII.

## ΝΕΟ ΚΡΙΤΗΡΙΟ ΑΠΟΔΟΧΗΣ 2026-11-16: REAL-DATA DEMO

**Real-data pilot ownership:** Αρχική πηγή/review `trial_source_alignment.py`, `trial_source_weekday_projection.py`, `trial_demo_source_reconciliation.py`; εγκρίσεις/identities `trial_patient_registry_projection.py`, `trial_import_plan.py`; gated Excel writer/readback `trial_demo_native_seed.py`; granular operational reader `outpatient_schedule_source.py`; frontend `trial_demo_weekly_navigation.py`, `trial_demo_patient_exit_vba.py`, `trial_demo_active_census.py`; day-effective patient exits `patient_exit_source.py`; opt-in MASTER `master_projection.py`, `master_projection_writer.py`.

**Gap ownership:** F/U/V source semantics still not admitted; private current actual programme + review and eligible Excel shell absent; the live real-data workbook + 1:1 manual reconciliation not performed; user flows and native save/reopen not accepted; in-session canonical MASTER refresh not automatically wired. Never substitute H8 synthetic data for business evaluation. Full matrix and next gate: `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`.

Τελευταία ενημέρωση: 2026-11-12

## Σκοπός

Αυτό είναι ο ζωντανός χάρτης ownership της υλοποίησης. Για συνολική κατανόηση του project ξεκινάμε από `Documentation/ONBOARDING.md`. Για το ακριβές resume point χρησιμοποιούμε `Documentation/CURRENT_HANDOFF.md`.

## Κύρια ownership περιοχών

### Scheduling / business logic

- `Python/rehab_core/availability.py`: patient/provider availability.
- `Python/rehab_core/conflicts.py`: scheduling conflicts.
- `Python/rehab_core/replacements.py`: current single-session replacement candidates and validation.
- `Python/rehab_core/replacement_options.py`: feasible replacement times and recommendation ordering.
- `Python/rehab_core/replacement_policy.py`: provider policy + patient schedule filtering.
- `Python/rehab_core/provider_policy.py`: provider profiles, blocks, max-load and leadership overrides.
- `Python/rehab_core/daily_state.py`: effective daily session state.
- `Python/rehab_core/orphan_resolution.py`: orphan-resolution semantics.
- `Python/rehab_core/daily_statistics.py`: statistics from effective daily state.
- `Python/rehab_core/workload.py`: authoritative therapist/student load from materialized operational sessions; outpatient sessions follow the same counting path.

The global multi-absence replacement optimizer is still open work. The existing single-patient ranking path must not be treated as the final global optimizer.

### Authoritative replacement facts / restore

- `Python/rehab_excel/daily_replacement_facts.py`: fact persistence model.
- `Python/rehab_excel/daily_replacement_fact_batch.py`: validate-all-before-mutate batch semantics.
- `Python/tools/restore_daily_replacement_fact.py`: single fact restore tooling.
- `Python/tools/restore_daily_replacement_facts.py`: selected/all authoritative restore backend.
- `Python/rehab_excel/daily_change_ledger_vba.py`: ledger VBA base.
- `Python/rehab_excel/daily_restore_in_session_vba.py`: in-session restore bridge.
- `Python/tools/build_daily_restore_bridge_preview.py`: native smoke/preview builder.

Authoritative truth remains `DAILY_REPLACEMENT_FACTS`; `DAILY_CHANGE_LEDGER` is only a projection/audit view.

### Daily views and MASTER

- daily render / write planning lives under `Python/rehab_excel/`.
- MASTER projection metadata uses OOXML/package reads rather than loading/saving the production workbook with openpyxl.
- `THERAPIST_DAILY`, `MASTER_SCHEDULE` and `DAILY_STATISTICS` must derive from the same effective daily truth.
- `Python/rehab_excel/therapist_workload_strip_vba.py`: MASTER workload visibility; renders `Όνομα n/6` badges from occupied `THERAPIST_DAILY` timeslot cells with green/yellow-red capacity thresholds.

### Outpatient recurring scheduling

Backend / Excel bridge ownership:

- `Python/rehab_excel/outpatient_schedule_bridge.py`
- `Python/rehab_excel/outpatient_schedule_registration.py`
- `Python/rehab_excel/outpatient_schedule_source.py`: authoritative outpatient recurring reader and unified recurring schedule merge.
- `Python/rehab_excel/outpatient_schedule_form_vba.py`
- `Python/rehab_excel/outpatient_schedule_edit_vba.py`: internal edit preload / stable target handling; public edit entry routes through the manager.
- `Python/rehab_excel/outpatient_schedule_manager_vba.py`: business-facing list/selection UI, hidden row-to-target mapping and registration-menu routing.
- `Python/rehab_excel/outpatient_conflict_guard_vba.py`
- `Python/rehab_excel/outpatient_in_session_vba.py`: in-session authoritative row write plus current daily projection refresh.
- `Python/rehab_excel/outpatient_daily_edit_vba.py`: transitional edit hardening that removes stale plain/decorated outpatient render lines from affected old/destination daily cells before rendering once.
- `Python/rehab_excel/scheduling_conflict_popup_vba.py`
- `Python/tools/build_outpatient_recurring_smoke.py`: one-session combined installer including manager, daily-edit hardening and therapist workload strip.

Current status:

- CREATE native green,
- stable same-row edit semantics implemented,
- application-style manager native-visible and usable,
- duplicate edit projection bug hardened; latest fresh native smoke appears correct,
- workload-strip UI implemented in code and awaiting focused/native verification,
- old row-number InputBox flow is retired from the public edit entry,
- conflict native gate follows successful workload visibility verification.

`OUTPATIENT_SCHEDULE` is hidden/protected authoritative storage, not a normal user workspace.

Relevant regression ownership:

- `Tests/test_outpatient_daily_edit_vba.py`: stale/decorated-line cleanup patch and idempotence.
- `Tests/test_outpatient_schedule_source.py`: outpatient source semantics plus one-entry/one-session/one-timeslot workload contract.
- `Tests/test_therapist_workload_strip_vba.py`: workload-strip action, timeslot counting and color-threshold contract.
- `Tests/test_outpatient_recurring_smoke_builder.py`: installer ordering and native build phase coverage including workload strip installation.

### Registration / application shell

- registration and top-level menu/form code lives under `Python/rehab_excel/` VBA generator/installer modules.
- `frmRegistrationMenu` remains reachable from the MASTER `Διαχείριση Μητρώων` action.
- its outpatient action is patched by `outpatient_schedule_manager_vba.py` to open the business-facing schedule manager rather than the raw create form.
- the product UX must remain application-like: MASTER/top buttons/forms, not native Excel grid navigation.
- detailed UI/product design history remains in `UI_APPLICATION_ROADMAP.md`; current cross-project UX rules are summarized in `Documentation/ONBOARDING.md` and locked in `Documentation/DECISIONS.md`.

## Closed correctness milestones

- restart-safe authoritative reconstruction: native green,
- replacement lifecycle V87-V92: native green,
- per-row / selected / restore-all authoritative restore: native green,
- daily statistics reconstruction parity: native green,
- previous outpatient combined smoke install: native green,
- outpatient CREATE: native green.

## Deferred performance work

Latest direct backend profile: `16.509s`. End-user restore latency in the low-to-mid 20-second range is accepted for now. Do not resume optimization unless there is a regression or final-polish phase.

Profiler remains at `Python/tools/profile_daily_restore_backend.py`.

## Open implementation milestones

1. focused therapist-workload-strip source gate,
2. fresh native workload-strip visibility and color-threshold gate,
3. outpatient native conflict gate on a separate fresh smoke,
4. remaining outpatient daily styling cleanup,
5. broader scheduling/UI roadmap,
6. global multi-absence replacement optimizer,
7. persistent pool/robotic operating-hours configuration,
8. final restore-performance polish later.

## Business rules that must not regress

- daily facts are overlays, never destructive recurring edits,
- capacity is timeslot-based,
- existing displaced patients precede new patients,
- every orphan resolves explicitly,
- MASTER and THERAPIST_DAILY agree,
- old restore cannot overwrite newer fact history,
- bulk restore validates the whole requested set before mutation,
- baseline is immutable,
- no production/staging `.xlsm` save with openpyxl,
- outpatients count operationally but do not belong in inpatient planner UI,
- one outpatient recurring entry must materialize to one operational session and contribute once to therapist active-session/timeslot workload,
- therapist workload display uses occupied timeslots, not patient-line count,
- outpatients never receive robotic treatment,
- technical sheet names, worksheet rows and internal IDs are not normal end-user concepts.

## Navigation

Start: `Documentation/ONBOARDING.md`

Then use:

- exact resume: `Documentation/CURRENT_HANDOFF.md`
- durable rules: `Documentation/DECISIONS.md`
- milestone/risk state: `Documentation/PROJECT_STATUS.md`
- history: `CHANGELOG.md`

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.


## Current recurring treatment ownership — 2026-11-12

Canonical state and run order: `Documentation/AUTHORITATIVE_RESUME_20261112.md`.

- `Python/rehab_core/recurring_therapy_changes.py`: effective-dated base schedule projections, preflight, historical date preservation; `Tests/test_recurring_therapy_changes.py`.
- `Python/rehab_excel/recurring_change_preview.py`: read-only workbook base/weekly/outpatient merge, approved capacity map, SHA/stale-row validation; `Tests/test_recurring_change_preview.py`.
- `Python/rehab_excel/stable_recurring_ids.py`: explicit per-assignment stable identifier binder; `Tests/test_stable_recurring_ids.py`.
- `Python/rehab_excel/recurring_id_migration_plan.py`: synthetic-only pure immutable-ID creation plan, rejects duplicates/ambiguous sources; `Tests/test_recurring_id_migration_plan.py`.
- `Python/tools/install_synthetic_recurring_ids_native.py`: one-shot *synthetic-only* native COM migration to a new XLSM, independent readback; `Tests/test_install_synthetic_recurring_ids_native.py`.
- `Python/rehab_excel/recurring_change_event_source.py`: optional strictly validated **read-only** `RECURRING_CHANGE_EVENTS` event reader (EventID, AssignmentID, PatientID, EffectiveDate, NewTherapist, NewStartTime, NewDayPattern, Reason); `Tests/test_recurring_change_event_source.py`.
- `Python/tools/prepare_native_probe_copy.py`, `probe_native_excel_readonly.py`, `prove_synthetic_cc_native_write.py`: isolated native structural copy/open/save/reopen checks and matching `Tests/test_prepare_native_probe_copy.py`, `Tests/test_probe_native_excel_readonly.py`, `Tests/test_prove_synthetic_cc_native_write.py`.
- `Python/tools/preflight_trial_excel_sources.py`: private real-staging and CC read-only admission gate. Never upload actual patient files to GitHub.
- `.github/workflows/diomil-synthetic-safety.yml`: full Python and targeted synthetic safety matrices, Linux and Windows; mocks only, not installed Microsoft Excel.

Accepted native evidence: frozen isolated read-only A, synthetic Save/reopen A and 179 immutable-ID migration A. **No native committed recurring event or rollback**. No operational migration approved. Runbook and frozen filenames: `Documentation/AUTHORITATIVE_RESUME_20261112.md`. Every future implementation must update this map alongside handoff/status/tests.



### Read-only permanent event append planner (2026-11-12)

- `Python/rehab_excel/recurring_change_append_plan.py`: source-SHA-bound event append *proposal*, strict stable ID and existing-event uniqueness/chronology, delegates to `recurring_change_preview` for therapist/patient capacity, never writes XLSM.
- `Tests/test_recurring_change_append_plan.py`: synthetic source immutability, duplicate ID, backdate, stale row and concurrent-change guards.
- CI: `.github/workflows/diomil-synthetic-safety.yml` targeted Windows/Linux plus complete suite. Native Excel event append not implemented.


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


## 2026-11-15 — VBA inventory ownership reconciliation

The native H10 inventory is complete for three distinct VBA projects. The 30 named application components map to tracked generators/installers; 37 components remain unresolved (10 generic standard modules and 27 worksheet/document modules). The 66-component CC candidate lacks one zero-line document module whose identity must be resolved with Unicode-safe output.

Detailed component-to-generator/test ownership and the next read-only write-surface gate are recorded in:
- Documentation/H10_VBA_OWNERSHIP_AND_WRITE_SURFACE_MAP_20261115.md

Do not infer that generic modules or worksheet document modules are obsolete. Do not implement blanket sheet protection or build the final synthetic integration workbook until their event/write/protection surfaces are reconciled.


## 2026-11-15 — real-source demo intake and latest H10 gate

The expanded H10 gate is operator-accepted (45 focused; 1,856 full passed, 1 skipped, 26 warnings; 3/3 distinct native candidates read-only audited, hashes unchanged, no workbook modified). See the master delivery report and current handoff.

The upcoming source-to-demo work should start by reviewing:
- `Python/rehab_excel/trial_admission_gate.py`
- `Python/tools/preflight_trial_excel_sources.py`
- `Python/rehab_excel/trial_import_plan.py`
- `Python/rehab_excel/trial_weekly_projection.py`
- their focused tests and `Documentation/REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md`

Build a synthetic fixture matching the operator's A:T programme mapping plus U/V afternoon/pool fields. Keep real source workbooks private. The identity review is not an approval ledger; admission must remain fail-closed until explicit identity and conflict decisions.


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


## 2026-11-16 — F/U/V read-only gate, 1,991 tests GREEN

**Concrete new gate, not a synthetic replacement for the user's real-data DEMO:** `Python/rehab_excel/trial_special_fields_audit.py`, `Python/tools/audit_real_manual_specials.py`, `Tests/test_trial_special_fields_audit.py` provide the first **single-source-only** private read-only inspection of actual manual A:V columns **F** (reclining, known 30-minute ΦΘ-related operation), **U** (legacy afternoon, no auto-import) and **V** (morning pool). No review/clinical approval/clean XLSM required for this *diagnostic* first command. It checks original source sheet/header, SHA-256 before/after, cached and raw Excel formula values (never omit formula-only cells), missing-patient source rows, time parseability, all counts and no PII emitted. The tool **never creates sessions, guesses F/V weekdays or imports old U**. It returns 0 if F/U/V truly empty (not whole source approval), 3 for pending real special records, 2 on safety stop. Exact local command, interpretations and privacy gates: `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md`.

**Verified code evidence:** `95157af7dbb32c6cb8aca210c503ffbb7f63be17`, GitHub Actions `43804737124`, **1,991 passing pytest in both Windows and Linux**, all four full/synthetic CI jobs successful. This is not proof that the real private manual has been scanned or a working native Excel was built. **Real-data DEMO remains NOT DELIVERED.** F/V day/provider/duration-aware scheduling approval, U source scope and independently reviewed outpatient programme remain blockers. Next meaningful step is run the read-only audit against the operator-approved private CURRENT manual snapshot, capture only totals, finish explicit F/V mapping and standalone outpatient source, then eligible private native seed and full side-by-side weekday/operator acceptance. Frozen H2–H10 proof artifacts remain untouched.

## 2026-11-16 — Existing MASTER content and formatting reused in real-data DEMO

**User clarification confirmed in source, NOT a new MASTER design:** `Python/rehab_excel/layout_contract.py` defines MASTER headers, treatment locations (D ΦΘ, E ρομποτικό, F πισίνα, G ανακλινόμενο, H εργο, I λογο, J ΕΦΑ, K ψυχολόγοι, L απογευματινό, M κατάσταση). `patient_program_master_render_vba.py` already implements `RefreshMasterPatientProgram`, `BuildTreatmentDisplay`: the actual treatment cell content is **clock (top line), weekdays (next line), optional approved provider (third line)**, wrapped and centered; `master_visual_style.py` owns existing colors, dimensions, header/identity styling, clinic banners. `PATIENT_PLANNER` has each matching `<treatment>_Ώρα`, `<treatment>_Ημέρες` and optional provider. These contracts are preserved rather than re-created.

**Integration defect discovered/targeted:** `master_projection_writer.py` must first sort admitted patients; its original one-cell formula templates, when used by the trial seeder, do NOT guarantee the already established multiline treatment display. New `Python/rehab_excel/trial_demo_master_display_bridge.py` reopens ONLY the disposable private staged `NATIVE_DEMO_STAGE.xlsm`, **requires** the already-installed `PatientProgramMasterRender.RefreshMasterPatientProgram` VBA module, calls it for each imported inpatient after the existing MASTER projection + visual style, and saves. It **does not install a duplicate formatter, change cell layout, modify the operational original or touch H2-H10 proofs**. Independent read-only `verify_private_demo_master_display` checks exact time, weekdays, optional provider and Excel WrapText for each losslessly represented (patient, treatment) in the initial trial; mismatches or a missing established VBA macro **fail closed** before publishing. The existing `DEMO_WEEKLY_VIEW` still shows all weekday/time variants where the legacy single-cell representation cannot show them losslessly. Import of actual source F/V/U special treatment fields is a SEPARATE unresolved input-admission contract; **none of the already supported treatment categories is prohibited**.

**Proof scope:** `Tests/test_trial_demo_master_display_bridge.py` checks original VBA column placements, full content/format order, sample persisted readback, deliberate tampering, missing/stale macro and seeder call order. CI for this latest commit is tracked separately. Native Windows/COM real-data candidate has **not** yet been executed. Required next actions are exact real manual source F/V/outpatient approval, private native seed and operator comparison, not another MASTER redesign.


## 2026-11-16 — Real outpatient source audit reuses existing application schema

**Current verified architecture:** Normal patient registration/program editor already writes `PATIENT_PLANNER` and uses `PatientProgramMasterRender.RefreshMasterPatientProgram` for the existing MASTER cell content. The new bulk real-data DEMO seed is a *different* initial-population path; it does not imply the old MASTER renderer was broken. Do NOT rebuild established cell formatting, controls or clinical treatment categories.

**Next separate real-data source bridge:** `Python/rehab_excel/trial_private_outpatient_audit.py` and `Python/tools/audit_real_outpatient_source.py` inspect only a PRIVATE, institution-approved current outpatient export in `.xlsx` form. They require the already-supported `PATIENTS` and `OUTPATIENT_SCHEDULE` sheets and their existing exact leading column headers. Reuse `read_patient_registry`, `read_outpatient_schedule` and `parse_day_pattern` to check outpatient identity/type/name consistency, Mon-Fri days, times, provider for ΦΘ, no robotic outpatient, per-weekday duplicate bookings, no formula-only or orphan source rows and SHA-256 read-only preservation. CLI emits only aggregate counts; never names/IDs/paths. `Tests/test_trial_private_outpatient_audit.py` uses only fictional private fixtures. **This is a diagnostic admission precondition, not an outpatient import or proof of completeness**. Full native outbound bridging from this separate approved source into the same private XLSM remains PENDING. Do not infer outpatients from the existing inpatient A:V review or omit them from the final pilot. Full command and schema: `Documentation/REAL_OUTPATIENT_SOURCE_EXISTING_SCHEMA_AUDIT_20261116.md`.

**Real workbook and operator acceptance status remains unchanged:** no current actual source files available in GitHub, no private native Excel produced or verified, outpatient source admission/merged seed and F/U/V approvals still open, 1:1 real manual comparison pending. Previous 2,002 test green code checkpoint `3cfdf61b` with GitHub run `43805735643` establishes the earlier MASTER bulk-display integration, not outpatient native acceptance. Check the new outpatient audit CI separately. Frozen H2–H10 acceptance files stay immutable.
