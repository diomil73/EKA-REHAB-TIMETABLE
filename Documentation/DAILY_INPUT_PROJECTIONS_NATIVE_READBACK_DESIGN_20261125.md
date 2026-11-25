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
