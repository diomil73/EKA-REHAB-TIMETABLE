# Implementation Map

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
