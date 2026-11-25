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
