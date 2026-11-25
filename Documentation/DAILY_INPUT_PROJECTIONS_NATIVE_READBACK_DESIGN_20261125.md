# DAILY_INPUT PROJECTIONS native readback — MASTER diagnostic slice

Date: 2026-11-25  
Status: **PARTIAL NATIVE DIAGNOSTIC IMPLEMENTED; NOT A PHASE VERIFIER; NOT AUTHORIZING**

## 1. Why this slice exists

The PROJECTIONS_REOPEN evidence contract requires more than saved formula links or occupied-clock counts. The existing saved-OOXML MASTER comparator and THERAPIST_DAILY occupied-clock comparator are useful separate diagnostics, but neither establishes that native Excel evaluates the saved MASTER formulas to the independently expected identity values, and neither proves complete therapist assignment membership.

This increment adds the missing native Excel read-only check for a deliberately narrow MASTER identity slice. It does not claim to complete PROJECTIONS_REOPEN.

## 2. Added implementation

- `Python/rehab_excel/daily_input_master_projection_native.py`
- `Tests/test_daily_input_master_projection_native.py`

Public diagnostic: `verify_master_projection_native_readonly()`.

The caller must supply an immutable `tuple[MasterProjectionRow, ...]` from an independently validated preflight. The expectation must not be reconstructed from the MASTER_SCHEDULE sheet being checked.

The diagnostic:

1. Requires an existing `.xlsm` file and hashes its exact bytes.
2. Runs the existing `verify_master_projection_saved_readback()` first. This retains the separate OOXML formula-link/schema/order/tail proof; it is not duplicated or relabelled as native Excel.
3. Starts a fresh isolated Excel COM instance with `DispatchEx`.
4. Sets `AutomationSecurity = 3`, disables visibility, alerts, events and link updates before opening.
5. Opens the exact evidence workbook with `ReadOnly=True`, checks the returned workbook is actually read-only and records Excel Version.
6. Rechecks the MASTER header/schema and formula linkage for the room, patient identity and status cells, then compares Excel-evaluated room, patient display identity and status against the independent `MasterProjectionRow` facts. Where a responsible doctor is rendered as a literal multiline identity, it verifies that exact literal rather than expecting a formula.
7. Closes with `SaveChanges=False`, quits Excel in cleanup, rejects cleanup failures and verifies the workbook SHA-256 is unchanged.

The diagnostic emits only aggregate metadata and generic failure messages; it does not include patient names/IDs in exceptions.

## 3. Strict scope boundary

This native diagnostic checks only:
- MASTER room value;
- MASTER patient display identity, including the approved responsible-doctor literal;
- MASTER status value;
- corresponding formula linkage for the checked cells;
- the existing OOXML comparator's full canonical formula structure.

It does **not** yet verify:
- every formula's evaluated result across all MASTER treatment columns;
- infectious formatting or every cell style;
- the full THERAPIST_DAILY projection through native Excel;
- actual session-to-patient-to-therapist AssignmentID extraction from the saved workbook;
- exact shared-slot membership from a native saved-source reader;
- every affected therapist/day projection or full source-input/ledger linkage;
- a phase report, receipt, trusted attestation or transition bundle.

Consequently, `EvidenceKind.PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`. No verifier registry status was changed. No phase adapter was added. This diagnostic cannot be used to satisfy the PROJECTIONS_REOPEN obligation.

## 4. Synthetic CI evidence

GitHub Actions run `45657152050`, code/test HEAD `e7ce7526de2c000cff9410819267f73fdc6ffc79`, completed **SUCCESS** across all four jobs:

- Windows Python 3.11 full regression: **2848 passed, 35 warnings, 48.64s**.
- Ubuntu Python 3.11 full regression: **2848 passed, 34 warnings, 28.83s**.
- Windows and Ubuntu synthetic safety jobs: success.

The new fake-COM tests exercise exact identity values, formula mismatch, wrong room/name/status, literal doctor identity, non-read-only open, missing Excel version, cleanup on verifier failures, and no mutation of the synthetic workbook. The workflow is synthetic regression evidence. The injected fake COM does not execute Microsoft Excel and is not native Excel acceptance. CI warning totals are nonzero; do not claim zero warnings.

## 5. Next engineering work

1. Audit and define the independently sourced expected MASTER values for every business-bearing projected column, without deriving them from the saved MASTER.
2. Build the independent native THERAPIST_DAILY saved-source reader. It must reconstruct exact `TherapistAssignmentFact` tuples keyed by immutable AssignmentID, PatientID, treatment, provider, service date and exact-minute time from a trustworthy source, not infer IDs from display labels or occupied clocks.
3. Compare those native observed facts to independently validated preflight facts and approved shared-slot membership. Keep occupied-clock totals as a separate supplementary invariant.
4. Bind the full set of affected MASTER and therapist scopes to the exact DAILY_INPUT and authoritative ledger facts from the same operation/source/date.
5. Only after those independent verifiers exist, add a PROJECTIONS_REOPEN report adapter and synthetic tests. Keep it non-authorizing until independently trusted verifier/build-manifest attestation is available.

## 6. Safety status

No production workbook, VBA, UI, journal writer, receipt, Apply/Restore path or Gate 2 setting was changed. No actual Excel COM or real XLSM was opened in this work. No real patient/staff data was used. Gate 2 remains closed.

 
## 7. Source-chain audit for the next therapist AssignmentID verifier

A source-only review of the current branch established the following existing contracts. These are reusable mechanisms; do not replace them with a new parallel scheduling interpreter.

### Existing stable identity and session construction

- `stable_recurring_ids.py::bind_stable_assignment_ids()` binds legacy planner entries through the unique `(PatientID, Treatment)` key in `RECURRING_ASSIGNMENT_IDS`; granular weekly and outpatient entries read their stable ID from the exact source row's `AssignmentID` column. Row numbers locate source data only; they are not the identity.
- `reader.py::read_base_schedule()` reads PATIENT_PLANNER and overlays `PROGRAM_WEEKLY_SLOTS` for matching patient/treatment pairs. `outpatient_schedule_source.py::read_unified_base_schedule()` merges the inpatient and outpatient sources.
- `rehab_core/base_schedule.py::materialize_sessions_for_date()` constructs `session_id = "{base_entry_id}@{YYYY-MM-DD}"`. With stable IDs bound first, the session ID retains the immutable AssignmentID prefix and explicit service date.
- `daily_replacement_facts.py` stores the exact `TargetSessionId`, patient, original/effective provider, replacement date/time, provider kind and ACTIVE/RESTORED state. Active replacement overlays are validated against the exact materialized session.
- `rehab_core/daily_state.py::build_daily_session_states()` preserves the original session identity while deriving active, absent, cancelled, replaced or pending-therapist-absence state and its effective provider/time.

### Important limitation of the rendered therapist view

The current `THERAPIST_DAILY` workload strip uses provider headers in rows 1 and 11, time labels in column A, and patient occupancy text in provider columns B:J across rows 2–8 and 12–18. The VBA workload function counts distinct occupied provider clocks; it does not persist AssignmentID in the visible grid. The existing `daily_input_therapist_assignment_membership.py` comparator explicitly requires independently sourced observed AssignmentID facts and forbids inferring IDs from therapist labels or clock counts.

Therefore the next native reader must not pretend that the display itself contains an AssignmentID. It must:
1. Capture the authoritative source sheets and the therapist display through the same fresh native Excel COM instance, with formulas/errors and finite row/column bounds checked.
2. Reuse the existing source readers, stable-ID binder, date-effective schedule and daily-state/replacement logic against the native-captured values, preserving original worksheet row identities for weekly/outpatient AssignmentID lookup.
3. Derive the expected effective therapist assignments from those independently validated source facts.
4. Resolve each native THERAPIST_DAILY patient/provider/time occupancy to exactly one expected assignment. A missing, duplicate or ambiguous join must fail closed; never guess a treatment or ID from display order.
5. Compare the resulting observed AssignmentID facts to the independent preflight tuple using `verify_therapist_assignment_memberships()`, including exact approved shared-slot membership.
6. Keep student replacements distinct from therapist assignments; do not count a student-provider cell as a therapist assignment.
7. Verify that all affected day/provider cells are covered and that the native sheet contains no extra occupancy, orphan rows or duplicate patients/times.

### Unresolved implementation questions to settle by source audit, not assumption

- Exact display-label transformation for outpatient names and any other presentation prefixes must be read from the actual daily write-plan builder before designing the join.
- The capture must preserve exact row numbers in `PROGRAM_WEEKLY_SLOTS` and `OUTPATIENT_SCHEDULE`, since the stable-ID binder uses those rows only to retrieve the immutable ID.
- The native reader must include the exact effective schedule semantics for PATIENT_EXITS and recurring change events, and must validate daily input against the authoritative ledger plus active replacement facts for the same date.
- Empty/zero-provider-day behavior and student replacement display must be explicit, not inferred from a nonempty-only unit comparator.
- The complete verifier must reconcile the same operation ID, source workbook digest, service date and preflight facts across DAILY_INPUT, ledger, MASTER and therapist projection. The existing single-sheet comparators remain supplementary checks, not substitutes.

This audit identifies a plausible reuse path, not a completed verifier. No new therapist reader or phase adapter is claimed by this section.

 
### 7.1 Exact patient-cell rendering contract inspected

The unresolved display-label question has now been traced to `patient_centric_preview.py`:

- A single patient/assignment in a provider-time cell renders the patient's display name.
- When different patients share the provider-time cell, each patient line receives its recurring day pattern in brackets.
- When the same patient has multiple assignments at the same provider/time, the patient name appears once and the following detail line lists each day pattern and treatment label.
- A temporary replacement keeps the original patient line styled/struck as appropriate and adds a source note `→ provider time`; the destination cell includes the patient name plus the temporary weekday/date label.
- Student replacement provider labels come from `student_display_state()`; they must remain a separate provider kind, not be counted as a therapist.
- Outpatient distinction is applied as cell presentation (fill) by `outpatient_presentation.py`; it does not add an outpatient prefix to the patient text in this renderer.

This means a future verifier should not invent a parser for loosely formatted patient names. It should reuse the exact established rendering functions/semantics to derive expected cell text from independently sourced daily states, then compare the native COM-visible cell value exactly. It must additionally reject ambiguous duplicate display names where the rendered cell cannot distinguish different PatientIDs, and it must compare the full set of affected provider/time cells so an unpatched stale cell cannot pass by omission.

This closes the source-audit question about patient-label transformation. It does not create the full expected-grid renderer or a native therapist verifier; those remain the next implementation.

 
## 8. Native THERAPIST_DAILY full-grid cell readback primitive added

Added:
- `Python/rehab_excel/daily_input_therapist_projection_native.py`
- `Tests/test_daily_input_therapist_projection_native.py`

Public diagnostic: `verify_therapist_daily_native_readonly()`.

It requires:
- exactly 18 independently supplied provider headers, ordered B:J in row 1 then B:J in row 11;
- an immutable expectation for **every one of the 126 patient-grid cells** in B2:J8 and B12:J18, including explicit blank expectations.

It opens the exact saved XLSM through a fresh macro-disabled Excel COM instance, requires ReadOnly, checks both provider header blocks and the fixed native time grid against the existing layout contract, rejects formula/error/non-text occupancy cells, compares every one of the 126 Excel-visible cell values exactly to the independent expected render map, captures Excel Version, closes without saving, quits and verifies unchanged workbook bytes. Cleanup failures fail closed. The reader does not infer AssignmentIDs from display text.

### CI

GitHub Actions run `45658828780`, code/test HEAD `71ff839699782615e9968d4213c9b8d42ea90bc6`, completed **SUCCESS** across all four jobs:
- Windows Python 3.11: **2859 passed, 32 warnings, 44.28s**.
- Ubuntu Python 3.11: **2859 passed, 35 warnings, 19.52s**.
- Both synthetic safety jobs passed.

The new tests use fake COM and synthetic workbooks. No Microsoft Excel process or actual workbook was opened. Warnings are nonzero.

### Boundary and next step

This is a native full-grid comparison primitive, **not a complete therapist assignment verifier**. It does not itself generate the expected 126-cell grid, independently reconstruct observed AssignmentIDs, prove exact shared-slot membership, or link the therapist grid to the same operation-bound DAILY_INPUT/ledger facts. It has no phase report adapter and does not change the verifier registry. `PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`.

Next engineering increment: build the complete expected-grid generator by reusing `read_unified_base_schedule()`, stable AssignmentID binding, date-effective entries, `materialize_sessions_for_date()`, the daily input/ledger, active replacement facts, `build_daily_session_states()` and the exact `patient_centric_preview` rendering semantics. It must provide all 126 cells, not just changed cells, and independently source the provider headers. Then join the full expected render map to immutable AssignmentID facts, rejecting duplicate display-name ambiguity, and bind all affected views to the same operation/source/date. Only after that should the PROJECTIONS_REOPEN report adapter be considered.

No production workbook, VBA, UI, journal writer, receipt, Apply/Restore or Gate 2 setting changed. Gate 2 remains closed.

 
## 9. Complete expected-grid builder added — source facts remain caller-supplied

Added:
- `Python/rehab_excel/daily_input_therapist_projection_expected.py`
- `Tests/test_daily_input_therapist_projection_expected.py`

Public builder: `build_therapist_daily_expected_grid()`.

It requires immutable tuples for:
- stable-ID-bound base schedule entries;
- materialized sessions for one explicit service date;
- exactly one DailySessionState per materialized session;
- the independent patient registry;
- optional student registry and provider label mapping;
- all 18 expected provider headers from an independent roster/layout contract.

The builder verifies stable AssignmentIDs, exact session/state/base-entry identity linkage, date consistency, unique PatientIDs and StudentIDs, provider-header agreement, and no StudentID/TherapistID collision. It constructs a complete 126-cell expectation, initially blank, fills every recurring provider/time group using the existing `render_base_slot()` semantics, then overlays the exact value patches from `build_patient_centric_preview_plan()`. It rejects ambiguous duplicate patient display names within a visible cell and any patch outside the confirmed grid.

It separately derives `TherapistAssignmentFact` expectations from active states and replacements to therapist providers. Patient absences, cancellations, unresolved therapist absences and student replacements are not counted as therapist assignments. The returned expected assignments remain preflight expectations, not observed native facts.

### Student provider rendering defect found and corrected

The first synthetic student-replacement case exposed a real issue in the existing patient-centric renderer: the final patch pass stored provider/time keys without provider kind, then tried to resolve a student destination as a therapist. Updated `patient_centric_preview._provider_cell()` to resolve a known StudentID through its student label even in that final pass. Also replaced stale `student.student_number` references in `patient_centric_preview.py` and `outpatient_presentation.py` with the canonical `Student.student_id` field.

An intermediate full CI run failed only the new student-replacement expected-grid test (2864 passed, 1 failed). The renderer fix, StudentID collision guard and dedicated collision test were added; the final complete run below is green.

### CI

GitHub Actions run `45660810319`, code/test HEAD `5a91058f0c9f78f63c19a70366987d8685796af7`, completed **SUCCESS** across all four jobs:
- Windows Python 3.11: **2866 passed, 36 warnings, 42.21s**.
- Ubuntu Python 3.11: **2866 passed, 32 warnings, 23.47s**.
- Both synthetic safety jobs passed.

The new tests use synthetic workbooks and fake COM only. No actual Microsoft Excel instance or real XLSM was opened.

### Boundary

The builder does not itself load or authenticate the current workbook's full source data, stable-ID registry, PATIENT_EXITS, recurring change events, DAILY_INPUT, authoritative ledger or active replacement facts. Those must be independently validated by its caller from the same exact source/operation/date. It also does not convert the native cell comparison into independently observed AssignmentID facts or a trusted phase report.

The native full-grid reader and this builder are now separate compatible primitives, but they are not yet joined by a native composite diagnostic/report adapter. `PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`; no verifier registry status changed. Gate 2 remains closed.

Next: build a single non-authorizing composite diagnostic that binds the exact source workbook hash, preflight facts, expected grid, native COM full-grid result and expected AssignmentID membership. Then add an independent saved-source/native capture path for the source sheets and ledger, preserving stable weekly/outpatient source row indices. Only after this is independently tested should a PROJECTIONS_REOPEN report adapter be considered.

 
## 10. Composite DAILY_INPUT / MASTER / THERAPIST_DAILY native readback binding

Added:
- `Python/rehab_excel/daily_input_projection_native.py`
- `Tests/test_daily_input_projection_native.py`

Public diagnostic: `verify_daily_input_projection_native_readonly()`.

This composes three separately produced results over one exact saved XLSM:
1. a previously successful `InputLedgerNativeReadback` for DAILY_INPUT and DAILY_CHANGE_LEDGER;
2. a fresh native MASTER readback against immutable `MasterProjectionRow` facts;
3. a fresh native THERAPIST_DAILY full-grid readback against the complete expected 126-cell grid.

It now requires an explicit SHA-256 for the independent MASTER preflight facts. It checks that this digest, the prior INPUT_AND_LEDGER result, the expected-grid source digest, both fresh native readbacks and the workbook bytes before/after composition all identify the same exact XLSM. It also binds one service date, checks the MASTER row count, the exact 18-header/126-cell therapist scope, the occupied-cell count, ledger/source counts and non-empty native Excel versions. Boolean-as-integer and missing/invalid metadata fail closed. The expected-grid builder now hashes the source before and after construction and carries that exact digest in its immutable result.

### CI

GitHub Actions run `45662463260`, code/test HEAD `6b55d2242b7ba1d3ec01ce01f5affa235717f94c`, completed **SUCCESS** across all four jobs:
- Windows Python 3.11: **2877 passed, 33 warnings, 45.11s**.
- Ubuntu Python 3.11: **2877 passed, 35 warnings, 18.84s**.
- Both synthetic safety jobs passed.

Tests inject fake COM readback results and exercise matching source/date, cross-file digest mismatch, date mismatch, incomplete grid counts, invalid metadata types and workbook mutation between sub-readbacks. No real Microsoft Excel instance or real XLSM was opened.

### Strict boundary

This is a **non-authorizing consistency composition**, not a PROJECTIONS_REOPEN phase adapter. It does not execute INPUT_AND_LEDGER itself; it accepts a prior readback object and checks its bindings. It does not authenticate caller-supplied result objects, prove their verifier provenance, bind an operation ID, create a report/receipt or establish trusted attestation. The expected MASTER and therapist facts still come from caller-supplied preflight; the builder does not independently capture/validate the complete source sheets, exits/recurring events, DAILY_INPUT, ledger and replacement store. Native MASTER evaluated-value scope remains limited to room, patient identity and status; full MASTER treatment-value/style verification remains open. AssignmentID facts in the expected grid are not independently extracted from the display cells; exact native text comparison is not itself a signed observed-identity attestation.

Therefore `PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`. No verifier registry status changed. Gate 2 remains closed.

### Next engineering step

Build a separate macro-disabled native COM source-capture/readback for the exact source sheets and operation facts. Preserve source row positions needed by the stable-ID binder, validate PATIENTS/PATIENT_PLANNER/PROGRAM_WEEKLY_SLOTS/OUTPATIENT_SCHEDULE/RECURRING_ASSIGNMENT_IDS, exits and recurring changes, DAILY_INPUT/ledger and active replacement facts, and feed only that independently captured result into the expected-grid builder. The capture must reject formula/error cells or unsupported formula semantics, enforce fixed bounds, preserve source hashes, and fail on missing/duplicate/ambiguous identities. Then expand MASTER native evaluated-value checks and only later consider a report adapter.

 
## 11. Blank THERAPIST_DAILY provider columns — layout compatibility hardening

A review of the existing OOXML therapist projection contract showed that unused provider header slots may be blank; the existing comparator rejects occupancy under a blank header but does not require all 18 header cells to be populated. The new native reader and expected-grid builder were corrected accordingly:
- expected provider headers still contain exactly 18 positions (B:J in both blocks), but an unused position may be an empty string;
- duplicate non-empty provider names are rejected within each block;
- a blank header must remain blank in native Excel;
- every patient cell under a blank header must be empty;
- all 126 patient cells and all 18 header positions are still inspected.

The first CI attempt exposed a test/diagnostic counter bug: the loop index used to select a cell's header overwrote the count of inspected headers. Corrected it by using a separate `cell_header_index`; the diagnostic now reports all 18 header positions checked.

GitHub Actions run `45663881144`, code/test HEAD `46d4023b0ba8e5c65510fd1ddc5bb249fdadb3ce`, completed **SUCCESS** across all four jobs:
- Windows Python 3.11: **2880 passed, 34 warnings, 60.57s**.
- Ubuntu Python 3.11: **2880 passed, 32 warnings, 27.13s**.
- Both synthetic safety jobs passed.

The added tests cover a valid unused blank provider column and reject occupancy under a blank header. Tests use synthetic workbooks/fake COM; no real Excel instance was opened. This is layout-compatibility hardening only; no verifier registry status changed. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED and Gate 2 remains closed.

 
## 12. Independent source preflight assembly and one-call native diagnostic

Added:
- `Python/rehab_excel/daily_input_projection_preflight.py`
- `Python/rehab_excel/daily_input_projection_preflight_native.py`
- `Python/tools/diagnose_daily_input_projection_native.py`
- `Tests/test_daily_input_projection_preflight.py`
- `Tests/test_daily_input_projection_preflight_native.py`
- `Tests/test_diagnose_daily_input_projection_native.py`

The preflight assembles existing canonical readers for one explicit service date and programme baseline date:
1. PATIENTS identity registry and SETTINGS provider roster;
2. unified inpatient/outpatient recurring programme;
3. stable recurring AssignmentIDs and recurring effective-date changes;
4. patient exits and active recurring entries;
5. dated sessions and effective weekly patient-slot validation;
6. DAILY_INPUT source changes;
7. active replacement facts and resulting daily session states;
8. STUDENTS and provider roster validation;
9. expected THERAPIST_DAILY 18-header / 126-cell grid and MASTER projection;
10. SHA-256 before/after the complete preflight.

The programme baseline date is deliberately mandatory and never inferred from filenames/current date. Empty/no-session days and empty DAILY_INPUT change sets currently fail closed. The preflight rejects missing/duplicate patient identities, unstable/duplicate recurring identities, unknown patient/provider references, date mismatch and workbook mutation.

Multiple patients assigned to the same provider clock require an explicit exact-membership approval tuple. The preflight does not infer approval from the collision. This tuple is only caller-supplied prior-decision data: it is immutable and exact-membership checked, but this implementation does **not** authenticate who approved it, bind it to an operation ID, or prove its source provenance. The CLI supplies no approvals and therefore fails closed on shared clocks until an independently trusted approval source is integrated.

The one-call native wrapper checks that the preflight and expected therapist grid refer to the exact source workbook hash/date, performs a native INPUT_AND_LEDGER readback, binds its exact source/active counts to the preflight changes, then runs the composite native MASTER and THERAPIST_DAILY readbacks. It checks the final workbook hash again. The CLI is Windows-only, macro-disabled/read-only through the underlying COM readers, requires explicit `--service-date` and `--baseline-date`, prints aggregate counts/hash/version only, and suppresses exception text that could expose patient/provider data.

Latest completed full CI before this documentation-only wording correction: run `45667684716`, HEAD `36e7e1d3646d378a49c5acfe92b80fd2acac6ca2`, SUCCESS across all four jobs:
- Windows Python 3.11: **2897 passed, 29 warnings, 48.44s**.
- Ubuntu Python 3.11: **2897 passed, 32 warnings, 19.99s**.
- Both synthetic safety jobs passed.

Synthetic tests cover preflight composition, unstable identities, invalid dates/providers, source mutation, shared-clock approval exactness, wrapper call sequence, hash/date/count drift, and safe aggregate-only CLI behavior. They use fake readers/COM and synthetic bytes. **No actual Microsoft Excel instance or real operational XLSM was opened.**

### Strict status boundary

This is a composed, non-authorizing diagnostic path. It is not yet a trusted source-capture attestation or PROJECTIONS_REOPEN phase adapter. Existing source readers are composed, not replaced by a single independently audited native COM capture of every source sheet. The approval tuple is not authenticated. Evidence objects are Python dataclasses and can be caller-constructed; there is no protected verifier registry/provenance token, operation ID, signed report or receipt. MASTER native evaluated-value scope remains limited to room, patient identity and status; treatment-value/style/package coverage remains open. Assignment identity in expected facts is not a trusted native observed-ID attestation. The CLI deliberately does not accept shared-slot approvals.

No verifier registry status changed. `PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`; Gate 2 remains closed.

### Next work

1. Independently capture and validate all source facts through native read-only Excel, with fixed worksheet/range bounds and exact sheet identities.
2. Define a trusted, operation-bound source for shared-slot approvals; never infer approval from observed duplicate clocks.
3. Extend native MASTER evaluated-value/style/package checks to the complete business contract.
4. Bind independently observed therapist assignment identity and all projection evidence to an operation ID.
5. Only then design the phase report/receipt adapter and operational real-XLSM smoke test.

 
## 13. Native literal source-sheet capture

Added:
- `Python/rehab_excel/daily_input_source_native_capture.py`
- `Tests/test_daily_input_source_native_capture.py`
- `Tests/test_package_probe_typed_cells.py`

The capture opens a separate native Excel COM session with macros disabled, events disabled, links not updated and the XLSM opened ReadOnly. It checks every source sheet present from this fixed allowlist:
`PATIENTS`, `PATIENT_PLANNER`, `SETTINGS`, `PROGRAM_WEEKLY_SLOTS`, `OUTPATIENT_SCHEDULE`, `RECURRING_ASSIGNMENT_IDS`, `PATIENT_EXITS`, `RECURRING_CHANGE_EVENTS`, `STUDENTS`, `DAILY_INPUT`, `DAILY_CHANGE_LEDGER`, `DAILY_REPLACEMENT_FACTS`.
The required minimum is PATIENTS, PATIENT_PLANNER, SETTINGS, DAILY_INPUT and DAILY_CHANGE_LEDGER; optional extensions are checked whenever present.

For each selected sheet it:
- rejects stored Excel error cells and all unsupported formulas;
- compares typed OOXML literal values with native COM values without collapsing text identifiers into numeric values;
- enforces per-sheet fixed row/column bounds and requires native UsedRange to cover all saved populated/formula cells;
- closes Excel without saving and requires the exact workbook SHA-256 to remain unchanged.

**PATIENT_PLANNER formula exception:** the established A:E identity bridge may contain either a direct same-row link to PATIENTS or the exact blank-guarded formula emitted by `planner_identity_formulas()`: `IF(PATIENTS!C{row}="","",PATIENTS!{column}{row})`. Only columns A:E and same-row references are supported. Native Excel must report the same formula and its evaluated result must match the corresponding literal PATIENTS cell. For the guarded form, Excel's exact-empty comparison is preserved: a genuinely empty PATIENTS name produces an empty string, while whitespace is not treated as empty. All other source formulas and stored error cells fail closed. Literal and supported formula counts are reported separately.

The one-call projection wrapper runs source capture before native INPUT_AND_LEDGER, then MASTER and full THERAPIST_DAILY readbacks. It validates source-capture result type, exact workbook hash, non-empty Excel version, unique sheet names, all five required sheets and nonnegative formula count. The CLI reports aggregate source sheet/literal/formula counts and source Excel version; no patient/provider cell values are printed.

CI run `45679125582`, code/test HEAD `2b9c6294232e476216f8438e2d1880797a539b0d`, SUCCESS:
- Windows Python 3.11: **2922 passed, 30 warnings, 42.15s**.
- Ubuntu Python 3.11: **2922 passed, 31 warnings, 29.86s**.
- Both synthetic safety jobs passed.

Tests cover supported direct and blank-guarded planner links, all five production identity formulas, formula result comparison to PATIENTS, exact blank-guard behavior, real OOXML package parsing, unsupported formula rejection, literal type preservation, formula/error rejection, fixed UsedRange bounds, source hash immutability, mandatory sheet scope and aggregate-only CLI output. Fake COM/synthetic only. No real Excel or operational XLSM was opened.

### Strict boundary remains

This verifies saved literal source-cell agreement and a deliberately narrow PATIENT_PLANNER formula contract. It does not attest source provenance or authorization, inspect styles/layout/package resources, authenticate shared-clock approval tuples, or bind an operation ID. The source preflight remains composed from canonical OOXML readers. Python evidence objects remain caller-constructible. No signed verifier attestation, report/receipt or phase adapter exists.

`PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`; Gate 2 remains closed.

## 14. Expanded native MASTER formula/evaluated-value coverage

The native MASTER reader inspects the complete 10-column canonical source-formula block plus the status formula for every projected patient row. It checks each formula against its exact PATIENT_PLANNER source-column/row link, reads all 11 Excel-evaluated cells, rejects formula errors and compares every formula result to the corresponding literal PATIENT_PLANNER source cell. Those source cells must not contain formulas. The source capture separately binds PATIENT_PLANNER literals/formulas to the supported source contract. The doctor-identity cell is the deliberate exception: when a responsible doctor is present, column C must be a literal and is compared to the independently expected patient/doctor text.

In addition to source equality, native evaluated values are compared to independent preflight facts for:
- infectious flag in MASTER column A;
- room in column B;
- patient/doctor identity in column C;
- status in the schema-selected status column.

The result records formula/evaluated-cell counts. The composite requires exact expected counts: 11 evaluated cells per projected row; 11 formula cells where doctor identity is formula-driven, or 10 where the doctor identity is literal. This prevents a partial native MASTER scan from being reported as complete.

Synthetic fake-COM tests cover all formula links, formula-to-source value mismatches, infectious mismatch, room/identity/status mismatch, missing/wrong formulas, formula error values, source formula rejection, literal doctor identity and composite formula/evaluated count enforcement. No real Excel instance or operational XLSM was opened.

### Remaining MASTER boundary

The seven other treatment-related formula outputs are compared to their PATIENT_PLANNER source cells, but are not yet compared to independently derived expected treatment/business facts. This does not verify full treatment semantics, formatting, styles, borders, row heights, conditional formatting or package resources. No operation ID, trusted verifier attestation, phase report or receipt is created. Gate 2 remains closed.

## 15. Explicit MASTER expanded-schema K/L boundary

The 13-column MASTER schema is `A:J` canonical formula projection, `K` psychologist, `L` afternoon programme, `M` status. The native MASTER verifier currently checks A:J and M only. This is not an accidental assumption that K/L are canonical formulas: `daily_input_master_projection_readback.py` explicitly leaves them outside that formula contract.

The current `MasterProjectionWriter` clears `A:P` for each projection row, then writes only the 10 canonical formula columns and status. Its existing expanded-schema test explicitly asserts that K2 and L2 are blank after projection. The separate VBA patient-program renderer can repopulate K (psychology) when a patient programme is edited, but it does not populate L. Afternoon physiotherapy is stored separately in `AFTERNOON_PHYSIO_SLOTS` and exposed through the read-only patient-form panel; no Master L renderer is part of this current projection path.

Therefore:
- K/L are deliberately outside the present DAILY_INPUT native projection proof;
- a green A:J/M readback must not be interpreted as proof that psychology or afternoon programme content is present or correct;
- any future decision to populate K/L in MASTER needs a separate explicit source-to-display contract, writer, native readback and tests before it can be included in the phase proof.

No change was made to this behavior. This is a documented scope boundary, not a claim that K/L are correct or complete.

## 16. Native MASTER verification of the expanded 13-column schema

Added fake-COM coverage for the expanded MASTER schema with K=`ΨΥΧΟΛΟΓΟΙ`, L=`ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ`, M=`Κατάσταση`. The native reader correctly identifies status column M and verifies the same A:J/M formula and evaluated-value contract. This test does not certify K/L content; those columns remain explicitly outside scope under §15.

CI run `45679125582`, HEAD `2b9c6294232e476216f8438e2d1880797a539b0d`, SUCCESS: Windows **2922 passed / 30 warnings / 42.15s**; Ubuntu **2922 passed / 31 warnings / 29.86s**; both synthetic safety jobs passed.
