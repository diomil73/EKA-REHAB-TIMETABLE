# DIOMIL practice DEMO: safe native seeding and operational acceptance

## 2026-11-16 — Executable private real-source DEMO preflight

**New executable interface:** `Python/tools/prepare_real_data_practice_demo.py` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Default operation is **read-only full real-source validation** for an authorized current A:V manual XLSX, confirmed review XLSX, entirely empty application-ready VBA XLSM and independently approved private `clinical_approvals.json`. Uses the strict production source alignment, all-patient approved ID check, original weekday/time/provider reconciliation, unresolved F/U/V stop gates, existing native shell checks and SHA-256 stability of the clinical-approval JSON. Every input must remain off Git checkout. No real patient name, ID, room, doctor, MRN, source path or clinical exception content is emitted to stdout/stderr on handled errors. `--build --output <PRIVATE_*_DEMO_*.xlsm>` is an *explicit* Windows + native Excel/COM, exclusive-new-file second step and repeats the same approval gates. **No real source file has been run; no real-data XLSM exists, and end-to-end practical acceptance is still pending.** Previous synthetic/H8 workbook results do not count as the operator's comparison. First actual action is an authorized private preflight with current manual and approved clinical review, not a random synthetic seed. Native failure or changed approvals after a native operation requires freezing the candidate, not treating it as accepted. See the dedicated operator CLI runbook for exact commands/JSON schema.

**Current blocker:** F/U/V and independently sourced outpatient bookings remain unavailable to the current *inpatient-only* A:V admission bridge. If occupied, the complete real-data source fails closed, never silently drops treatments. Do not alter frozen H2-H10 evidence. Update all living handoff/status/test and delivery documents again with actual source-specific counts and results **only after** the first private native execution, without publishing PII.

## ΕΠΙΧΕΙΡΗΣΙΑΚΗ ΕΝΗΜΕΡΩΣΗ 2026-11-16: ΠΡΑΓΜΑΤΙΚΟ DEMO

**Real-data only for the operator DEMO.** Synthetic data are allowed ONLY for CI/safety testing, not for clinical workflow comparison. The eligible private original A:V programme must be the **current manual snapshot** for the same date/week, plus approved review identities and independently approved clinical facts. Use a completely clean isolated VBA XLSM shell, all SHA/safety gates, COM write/readback and side-by-side per-patient/day/time/treatment/provider reconciliation. No identity/status guesses, no partial source admission, no weekend treatments, no public real-data uploads. Existing F/U/V (reclining 30 min, afternoon/pool) remain blocking until a correct representation is accepted. Real native output and operator acceptance NOT YET DONE; the seeder is internal code, not a user-ready one-click build. Detailed acceptance and issue handling: `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`. When a real DEMO is actually created, record its **non-identifying** evidence, native/save/reopen outcome, unresolved issues, SHA, exact user checklist results, and updated filenames/commit in all live handoffs.

Date: 2026-11-15
Branch: `work/afternoon-medical-scope-20261115`
**Current status: code bridge + synthetic tests only. No operator/native Excel acceptance yet.**

## Goal

Deliver **one separate macro-enabled workbook**, with the existing Diomil navigation and daily operational workflows. The first practical acceptance scenario must include initial patients/schedule, new inpatient admission, discharge, patient absence, therapist absence and replacement/restore, new therapist, new student, and save/reopen day. Analytics/H9/H10 visual refinements and fifteen-day progress may follow.

The current source/projection/native-seed modules form a necessary import foundation, but alone DO NOT certify that this interactive workbook is ready for exercise.

## Private data safety

The original programme A:V XLSX, the reviewed identities XLSX and the eligible empty XLSM remain on institution-controlled private Windows storage. **Never commit any workbook bytes, patient names, MRNs, room/doctor metadata, source-cell content or real schedules to the public GitHub repository.**

No source or frozen H10 candidate may be overwritten. Always use a fresh `*_DEMO_*.xlsm` output path outside the Git checkout. A failed native seed must not be retried on the same output. Source SHA256 checks are mandatory before and after any copy/COM write.

## Source and admission gates

1. Operator-confirmed A:V interpretation: no weekend treatments. A treatment time with blank days means exactly ΔΕ, ΤΡ, ΤΕ, ΠΕ, ΠΑ. No time means no slot; explicit day patterns preserved exactly.
2. The original source row, clinic label, patient name and ward/doctor context must match the reviewed workbook with unchanged SHA256. Blank column A stays blank, never inherited from the preceding row.
3. The source and review must contain explicit confirmed PatientIDs, resolved holds/overlaps, and exact shared provider slot approvals. Never promote suggestions automatically.
4. `project_trial_demo_sources` checks the original source again and proves the complete equality of treatment/day/time/provider/source-cell tuples before preparing PATIENTS and granular weekly rows. F reclining and U/V afternoon/pool remain **held** until their independent operational semantics are delivered. There must be no silent partial import.
5. Clinical classification, infectious status and room/doctor/status are independently approved; neither the original context text nor Excel formatting authorizes guessing.
6. The `empty_demo_shell=True` plan is allowed only when the existing workbook inspection finds exactly the `EMPTY_OR_MISSING_PATIENT_REGISTRY` issue: no patients, no history, valid VBA and core sheets. The native seed performs an additional stricter check for empty PATIENTS, treatment planner, weekly source, prior students and active operational sheets.

## Engineering entry points (not an operator command yet)

- `rehab_excel.trial_import_plan.prepare_trial_import_plan(review=..., intake=..., programme=..., empty_demo_shell=True)`
- `rehab_excel.trial_demo_source_projection.project_trial_demo_sources(plan, programme=..., review=..., intake=..., clinical_approvals=...)`
- `rehab_excel.trial_demo_native_seed.seed_trial_demo_workbook(plan, programme=..., review=..., intake=..., output=..., clinical_approvals=...)`

`clinical_approvals` is a **private in-memory** mapping keyed by (review clinic, original Excel row), with `ConfirmedClinicalFields` values. The interface is internal and is not yet a one-click operator workflow. There is no approved private data payload or clean source shell currently available through the GitHub-only execution environment.

The Windows COM backend writes only to the isolated private staged XLSM. It writes PatientID/room/name/infectious/status and optional type/MRN/doctor/context columns in PATIENTS; identical-row identity formulas in PATIENT_PLANNER; one source-bound row per scheduled treatment/day in PROGRAM_WEEKLY_SLOTS; and a detailed DEMO_WEEKLY_VIEW for non-compressible timings. It then performs independent readback, reads the weekly schedule through the real scheduling reader, rechecks inputs, and exclusively creates the final output. Native source visuals and VBA are not removed.

## Critical remaining UI gap

Existing MASTER/PATIENT_PLANNER presentation cannot losslessly represent all distinct treatment times on different weekdays in a single cell. The new DEMO_WEEKLY_VIEW shows the detailed programme, but existing central navigation and replacement screens still need integration/real user validation for those granular rows. A successfully seeded candidate **is not** the final accepted interactive DEMO until those workflows read the same appointments and display them correctly.

Likewise, the current F/U/V treatment holds block whole-source admission. They must either be implemented under separately accepted clinical rules or managed as explicitly unresolved records in a designed, visibly partial practice environment, never treated as secretly dropped sessions.

## Native operator acceptance checklist

Use a **new disposable private candidate**, preserve a frozen clean starting copy for repeatable practice, and only then test:
1. Launch, all forms/navigation/toolbar buttons, identity and clinic sections.
2. Original recurring weekday/time/provider schedule and detailed day view agree; no phantom Saturday/Sunday sessions; no double counting of shared slots.
3. Therapist absence triggers candidate replacement choices, capacity/skill restrictions and approval.
4. Patient absence, return and allowed daily restoration do not destroy immutable recurring assignments.
5. Add inpatient, discharge existing patient, add therapist and student with appropriate eligibility/dates; daily updates reflected in all connected screens.
6. Save/close/reopen through native Excel; verify workbook formulas, VBA project, programme, history and unchanged original source.
7. If any action fails, preserve the candidate and logs, do not force repairs on frozen sources or repeat unneeded old test gates.

**Acceptance wording:** A green Python CI run is code-level evidence only. A new native COM write/readback is a separate milestone; complete interactive operator DEMO acceptance is later still.


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


## 2026-11-16 — F/U/V read-only gate, 1,991 tests GREEN

**Concrete new gate, not a synthetic replacement for the user's real-data DEMO:** `Python/rehab_excel/trial_special_fields_audit.py`, `Python/tools/audit_real_manual_specials.py`, `Tests/test_trial_special_fields_audit.py` provide the first **single-source-only** private read-only inspection of actual manual A:V columns **F** (reclining, known 30-minute ΦΘ-related operation), **U** (legacy afternoon, no auto-import) and **V** (morning pool). No review/clinical approval/clean XLSM required for this *diagnostic* first command. It checks original source sheet/header, SHA-256 before/after, cached and raw Excel formula values (never omit formula-only cells), missing-patient source rows, time parseability, all counts and no PII emitted. The tool **never creates sessions, guesses F/V weekdays or imports old U**. It returns 0 if F/U/V truly empty (not whole source approval), 3 for pending real special records, 2 on safety stop. Exact local command, interpretations and privacy gates: `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md`.

**Verified code evidence:** `95157af7dbb32c6cb8aca210c503ffbb7f63be17`, GitHub Actions `43804737124`, **1,991 passing pytest in both Windows and Linux**, all four full/synthetic CI jobs successful. This is not proof that the real private manual has been scanned or a working native Excel was built. **Real-data DEMO remains NOT DELIVERED.** F/V day/provider/duration-aware scheduling approval, U source scope and independently reviewed outpatient programme remain blockers. Next meaningful step is run the read-only audit against the operator-approved private CURRENT manual snapshot, capture only totals, finish explicit F/V mapping and standalone outpatient source, then eligible private native seed and full side-by-side weekday/operator acceptance. Frozen H2–H10 proof artifacts remain untouched.
