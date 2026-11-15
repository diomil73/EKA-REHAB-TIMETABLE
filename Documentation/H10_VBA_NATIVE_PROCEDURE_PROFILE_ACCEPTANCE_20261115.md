# H10 — Native VBA procedure-profile acceptance and remaining behavioral boundary

Date: 2026-11-15
Branch: `work/afternoon-medical-scope-20261115`
Evidence set: `DIOMIL_VBA_PROFILE_BUNDLE_20261115_195014.txt`, `DIOMIL_VBA_PROFILES_20261115_195014.txt`, `DIOMIL_PROFILE_FULL_20261115_195014.txt`, `DIOMIL_PROFILE_FOCUSED_20261115_195014.txt`.

## Accepted gates

| Gate | Observed result | Acceptance scope |
|---|---|---|
| Focused audit tests | 27 passed in 1.71 s | Source regression only |
| Full source regression | 1,832 passed, 1 skipped, 26 warnings in 31.79 s | Source regression only, not native UI acceptance |
| Native read-only audit | Three distinct H10 candidates completed | Structural write-surface inspection |
| Workbook integrity | 3/3 `SOURCE HASH UNCHANGED: YES` and `NO WORKBOOK WAS MODIFIED` | No tested source-file mutation |
| Security | `ReadOnly: True`, macros and events disabled before/open during inspection | No macro execution during audit |

The one skipped test is `Tests/test_afternoon_sheet_preview.py` due to unavailable symlinks. The warnings include zipfile deallocator warnings and the openpyxl unsupported Data Validation extension. The earlier native VBA component-inventory run found that setting Excel Calculation to Manual was unavailable due to a COM error; the new audit does **not** claim Manual Calculation was applied.

Three distinct candidates inspected:
1. `UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm` (67 VBA components)
2. `UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm` (67 VBA components)
3. `UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm` (66 VBA components)

Each reports `Matched operations: 1569`, `VBA WRITE-SURFACE AUDIT COMPLETE`, unchanged source SHA-256 and no workbook mutation. The 1,569 findings are **static line-pattern matches**, not 1,569 confirmed runtime writes.

## Cross-candidate reconciliation

All three reports contain **44 nonempty VBA components** represented by `COMPONENT` or `NO_PATTERN_MATCH_COMPONENT` entries. Comparing their component metadata and matched operation/procedure records shows identical reports for all 44, excluding only the workbook header/name. The 23 or 22 other modules, depending on candidate, have zero code lines and therefore are not emitted by this write-surface report.

Eight components per candidate have no detected operation-pattern match. Their component SHA-256 values and procedure-profile hashes are identical across all three candidates. This is **direct code-body fingerprint evidence for these eight**, unlike the 36 components with detected operations for which the present report supplies only operation metadata, not full-module hashes.

### Previously unresolved native-only modules

| Module | Source-redacted evidence | Supported conclusion | Remaining limit |
|---|---|---|---|
| ThisWorkbook code-behind | 19 lines; `Workbook_Open`, `Workbook_Activate`, `Workbook_Deactivate`, `Workbook_SheetActivate`, `Workbook_BeforeClose`; each event body one line categorized as call. Open/Activate share identical body hashes; Deactivate/BeforeClose share identical body hashes; SheetActivate has a third body hash. | Five event hooks are present and their body fingerprints match across the three candidates. | Called procedures/side effects must be traced in tracked application-shell source or controlled native source review. |
| `PATIENT_PLANNER` sheet code-behind | 25 lines; `Worksheet_Change` with 23 body lines: 7 branches, 2 event/target references, 3 calls, 3 other; same fingerprint across all candidates. | Nontrivial event handler exists; its behavior cannot be declared read-only from absent worksheet-write patterns. | Exact call targets and event-trigger boundaries remain unproved. |
| `Module6` | 26 lines; `NormalizeTimeMS` (9 body lines, 5 branches, 2 date/time candidates), `NormalizePresenceMS` (10 body lines, 4 branches, 4 assignments, 1 other); same procedure hashes across all candidates. | Helpers normalize time/presence by naming and structural patterns. | Names/counts do not prove exact normalization rules, accepted values or caller set. |
| `Module5` | `CountOfLines=1` but no procedure, and `component_sha256=e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`, the SHA-256 of an empty UTF-8 string. | The native `CodeModule.Lines` output normalized to empty content for this module in every candidate. | Retain the module unchanged; do not delete it based on the profile. |

### Additional four no-pattern entry components

`modRegistrationMenu` (`ShowRegistrationMenu`), `modOutpatientSchedule` (`ShowOutpatientScheduleForm`), `modOutpatientScheduleManager` (`ShowOutpatientScheduleManager`) and `modSchedulingConflict` (`ConfirmTherapistDoubleBooking`) match tracked Python VBA generators as previously documented. They have unchanged full component hashes across the three candidates. Their source roles are UI/navigation and explicit confirmation, not direct worksheet writeback in those entry modules. Delegated paths must still be included in the ownership matrix.

`modDoubleBookingReview` is not no-pattern. Its 16 detected operations include external temporary file I/O, process launches, and workbook-close lifecycle; it remains a write/lifecycle owner.

## Status and next work

- **GREEN:** Focused tests, full regression, native component inventory, read-only write-surface audit, cross-candidate consistency of detected operations and fingerprints for no-pattern modules.
- **PARTIAL:** Source ownership/call graph, full semantics of native-only workbook and sheet event handlers, precise writer-to-sheet execution-path matrix.
- **NOT ACCEPTED:** Complete native, end-to-end integrated demo; importer/admission of the real patient treatment workbook; operational migration.

Recommended next autonomous source work: trace the application-shell and event-related callable names available in tracked Python-generated VBA source. If exact embedded event-call destinations are not accessible in tracked source, keep those as an explicit native inspection blocker rather than invent them. Do not change source workbooks or promote them to writable integration baselines. The next test workbook must be a newly named synthetic copy, after the separate controlled source-admission decision.

## Provenance and safety

The user-provided reports are local audit artifacts, not new Git-tracked production content. The acceptance is limited to report-observed results; no native Excel COM run occurred in this documentation task. No patient data, VBA source text, cell formulas or patient identities are reproduced in this register.


## Subsequent source-level ownership proof: ThisWorkbook events

The tracked generator `Python/rehab_excel/application_shell_vba.py` contains `_EVENT_SPECS` and `install_application_shell`, which install calls into the workbook's VBA document module.

We independently computed the SHA-256 of the exact one-line VBA body created by each generator entry (four spaces followed by the call name). Every result **matches the observed native profile hash** in all three candidates:

| Workbook event | Installed call | SHA-256 prefix |
|---|---|---|
| `Workbook_Open` | `EnterApplicationShell` | `587b673406a7` |
| `Workbook_Activate` | `EnterApplicationShell` | `587b673406a7` |
| `Workbook_Deactivate` | `ExitApplicationShell` | `c755c205e4b4` |
| `Workbook_SheetActivate` | `KeepApplicationShell` | `04ddae21bbda` |
| `Workbook_BeforeClose` | `ExitApplicationShell` | `c755c205e4b4` |

**ThisWorkbook event source ownership is now confirmed** for this one-line event-handler scope, not merely inferred from procedure names. The called routines are implemented in the tracked `modApplicationShell` generator: entering/preserving the application view, restoring user interface on deactivate/close, and retaining the sheet/navigation shell. Nested calls such as `ApplyMasterView` are still separately governed by their existing contracts.

A regression assertion was committed in `Tests/test_h10_native_workbook_event_hash_contract.py` to prevent generator event calls from drifting away from these native SHA-256 proofs. This new test has not yet been run on the operator machine; do not include it in the earlier 1,832-pass claim.

The remaining native-only behavior review narrows to `PATIENT_PLANNER.Worksheet_Change`, `Module6` time/presence helpers and the empty `Module5` module. Module5 is confirmed empty in the native code-module output, so there is no executable procedure to trace. The exact implementation and called routines of `PATIENT_PLANNER.Worksheet_Change` still require inspection without exposing patient data or modifying a workbook.
