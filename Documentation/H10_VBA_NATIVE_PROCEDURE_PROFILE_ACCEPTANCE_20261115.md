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


## Machine-checkable write-surface evidence reconciliation

Added `Python/tools/verify_h10_vba_profile_reports.py` and `Tests/test_verify_h10_vba_profile_reports.py`. These parse the sanitized UTF-16/UTF-8 auditor output and fail closed on absent read-only/macros-off/events-off/hash-unchanged/no-modification markers, unexpected workbooks/component counts, incomplete/truncated audit, missing per-component fingerprints, and cross-candidate operation/profile differences.

The uploaded `DIOMIL_VBA_PROFILES_20261115_195014.txt` was independently parsed for the following per-workbook categories (identical across all three distinct candidates):

| Static line-pattern category | Count |
|---|---:|
| worksheet_reference | 80 |
| cell_value_or_formula_access | 373 |
| cell_value_or_formula_write | 282 |
| range_or_structure_mutation | 111 |
| format_or_other_access_or_write | 464 |
| visibility_navigation_or_action_binding | 30 |
| external_process_launch | 28 |
| external_file_write_or_delete | 91 |
| external_file_read | 46 |
| workbook_save | 26 |
| workbook_close | 4 |
| protection_unprotect | 17 |
| protection_protect | 17 |
| **Total** | **1,569** |

All three report the exact same component-category counts and operation metadata for 44 nonempty VBA modules. These counts describe static pattern lines, not confirmed mutating instructions or operational call execution. Explicit worksheet targets in the report are only those syntactically named in a matched VBA line. They are *not* a complete variable/data-flow mapping.

High-priority static sources for human ownership review:
- `modTherapistHistoryReport`: 271 matched lines, including 74 direct Value/Formula assignment candidates, 22 range/structure candidates, and 10 protection calls.
- `modPatientRegistrationBridge`: 111 matched lines, including 42 direct assignment candidates across registration/update code.
- `modDailyChangeLedger`: 106 matched lines, including 29 direct assignments and restore/rollback operations.
- `modGlobalReplacementAction`: 96 matched lines, including 25 direct assignments, saves, protect/unprotect and external Python/file operations.
- Legacy `Module4`: 87 matched lines, including 18 direct assignments, 8 structure candidates and references to the replacement workflow.
- `modDoubleBookingReview`: 16 matches involving file I/O, Python execution and workbook lifecycle.

The new automated verifier and its synthetic unit tests were committed **after** the reported 1,832-pass source regression, so they are **not yet operator-tested**. The independent ad hoc inspection of the uploaded report is not a substitute for a passing repository test suite. Do not promote the release acceptance level because of this addition.

No clinical data, original VBA source code or workbook files were added to the GitHub repository.


## 2026-11-15 — Offline evidence verification and prioritized writer map

The uploaded `DIOMIL_VBA_PROFILES_20261115_195014.txt` is UTF-16LE with a BOM. A direct local invocation of the same parsing and comparison logic as `Python/tools/verify_h10_vba_profile_reports.py`, operating on that exact uploaded artifact, returned **GREEN**: three named candidates (67/67/66 total VBA components), 44 nonempty component records per candidate, exactly 1,569 operation-pattern rows per candidate and 12 no-pattern procedure profiles apiece. The component metadata, recorded operations and procedure-profile hashes all agree across the three. This verifies the report data and its parser logic; it does **not** substitute for running the committed repository pytest tests or for a new Windows Excel inspection.

Prioritized static ownership table for review (counts per candidate; all three agree):

| Module | Value/formula write candidates | Range/structure candidates | Protection calls | Important note |
|---|---:|---:|---:|---|
| `modTherapistHistoryReport` | 74 | 22 | 10 | High concentration of derived history/report changes |
| `modPatientRegistrationBridge` | 42 | 5 | 0 | Registry/planner and related views, explicit `PATIENTS` / `PATIENT_PLANNER` references |
| `modDailyChangeLedger` | 29 | 5 | 0 | Daily ledger/restore and authoritative save hooks |
| `modGlobalReplacementAction` | 25 | 10 | 4 | Replacement transactions, Python/file I/O and rollback |
| `modReplacementOperator` | 20 | 9 | 6 | Operator replacement decisions and protections |
| Legacy `Module4` | 18 | 8 | 0 | `REPLACEMENTS` / `REPLACEMENT_LOG` and planner logic |
| `modPatientHistoryReport` | 13 | 4 | 6 | History report and protection behavior |
| Legacy `Module2` | 11 | 10 | 0 | Therapist attendance and conflict refresh |
| `modOutpatientInSession` | 10 | 1 | 0 | Day and outpatient in-session state |
| Legacy `Module1` | 8 | 7 | 0 | Therapist daily presentation |
| `modDoubleBookingReview` | 0 | 0 | 0 | **Not read-only**: 5 process-launch and 6 file-write/delete candidates, plus workbook close |

The native pattern output names 13 explicit worksheet targets: `CONFLICT_LOG`, `DAILY_INPUT`, `FINALIZED_DAY_SNAPSHOTS`, `MASTER_SCHEDULE`, `NEW_PATIENT`, `OUTPATIENT_SCHEDULE`, `PATIENTS`, `PATIENT_PLANNER`, `REPLACEMENTS`, `REPLACEMENT_LOG`, `SETTINGS`, `THERAPIST_ATTENDANCE` and `THERAPIST_DAILY`. This list records **literal references in matched lines only**; it does not mean no other sheets are written through variables.

To make the remaining event/normalizer analysis more specific without exposing raw source, the read-only auditor now records `POSSIBLE_CALL_EDGE caller=... line=... callee=...` for the eight no-pattern components. The callable names are collected exclusively from procedure declarations in the inspected VBA project, and string literals/comments are suppressed during matching. **These are candidate identifier references, not proved runtime calls or control flow.** New synthetic tests were added; both the code and tests are pending the next normal operator regression/native gate. This does not change any workbook or Trust Center setting.

Next decision gate: reconcile `PATIENT_PLANNER.Worksheet_Change` candidate edges to tracked writer/refresh routines, examine `Module6` normalization callers, then design the fully synthetic integrated workbook builder in a brand-new path. The frozen H10 candidates remain immutable.


## 2026-11-15 — call-edge reconciliation now covered by tests

The targeted, source-redacted `POSSIBLE_CALL_EDGE` output was integrated into `Python/tools/verify_h10_vba_profile_reports.py`. The verifier now compares candidate references between all three distinct workbooks in addition to component metadata, operation matches and procedure fingerprints. New synthetic tests in `Tests/test_verify_h10_vba_profile_reports.py` exercise matching call edges and an intentionally divergent third-candidate edge. Quoted string literals and VBA comments are excluded during the native candidate extraction; the output is still only a *possible reference*, not proven runtime call execution.

**Latest change test status:** the new call-edge extraction, verifier extension and regression assertions have been committed but have not yet been executed on the operator's Windows source checkout. The previous 1,832-pass full suite is not proof for these new commits. The previously uploaded native profile report did not yet include call-edge records, so it cannot establish the unknown PATIENT_PLANNER event's actual call targets.

No additional Excel workbook was opened or saved during this GitHub-only update; previously frozen native evidence remains unchanged.


## 2026-11-15 — verifier hardening and one-command Windows call-edge gate

Changes implemented on `work/afternoon-medical-scope-20261115`:

1. Fixed a real parser defect: the `POSSIBLE_CALL_EDGE` regex accidentally escaped `\\s`, `\\S` and `\\d` twice, preventing correct recognition. The corrected expression has an explicit parsing regression test.
2. The native read-only auditor now searches **all components with existing operation matches** for potential references to `NormalizeTimeMS` and `NormalizePresenceMS`, not just the eight no-pattern components. Existing no-pattern call-edge records are retained. These are **candidate identifier references**, not proof of runtime invocation.
3. Corrected normalized no-pattern `component_sha256`: its line separator is now an actual LF, not a literal backslash plus `n`. **Previous native source-redacted component fingerprints are historical fingerprints from the old algorithm**; they are comparable within that original report, but must not be compared directly against newly generated component fingerprints. Per-procedure hashes were already normalized with LF and are not affected by this component-hash correction.
4. The verifier now refuses malformed operation/profile/call-edge records and compares each component's declared `matched_operations` count to the number of parsed operation rows, in addition to global counts, safety attestations and three-candidate equality. Regression cases for malformed input and mismatch were added.
5. Created `Scripts/run_h10_vba_call_edge_gate.ps1`. It enforces the correct branch, existence of three frozen reference XLSMs, fresh unique output logs, focused pytest, full pytest, read-only native auditor, 3/3 successful audit/hash/no-modification markers and the offline three-candidate verifier. It bundles results into a single upload file. The script never saves or edits a source XLSM and does not automatically retry failed candidates.

Archived uploaded report `DIOMIL_VBA_PROFILES_20261115_195014.txt` was inspected offline again: exactly three completed audit blocks, each with 1,569 parsed operation rows, 12 procedure profiles and eight no-pattern components. This is a read-only check of **old evidence**, not proof for the newer scanner/verifier or the new one-command gate.

**Validation status:** All newest source changes and PowerShell runbook are committed, but no post-change Windows PowerShell/pytest/native run has yet been observed. Do not describe the new version as accepted until the next operator report. The previously accepted 1,832-pass regression and three preserved source hashes remain valid for their earlier revision. The next required native evidence is not an integration installation; it is the targeted call-edge proof needed to resolve `PATIENT_PLANNER.Worksheet_Change` and `Module6` helper callers.

### Exact operator command when the next gate is ready

With Excel closed, from the repo directory:

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull failed" }
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\run_h10_vba_call_edge_gate.ps1
```

Upload the printed `DIOMIL_H10_CALL_EDGE_BUNDLE_*.txt`. The PowerShell script stops on failed source regression, absent candidates, incomplete native results, changed workbook fingerprints, or cross-candidate differences. It does not grant permission for a writable synthetic demo or for any real patient migration.

After successful evidence review, continue the new fully synthetic integrated fixture design, maintaining explicit separation from frozen native candidates and real admission paths.


## 2026-11-15 20:31 — H10 call-edge gate accepted; event/validator ownership narrowed

Primary operator artifact: `DIOMIL_H10_CALL_EDGE_BUNDLE_20261115_203111.txt`.

**Observed Windows results:**
- Focused tests: **38 passed in 1.94s**.
- Full repository regression: **1,854 passed, 1 skipped, 26 warnings in 31.08s**; no failures.
- The one skip remains `Tests/test_afternoon_sheet_preview.py:68`, because symlinks are unavailable in the Windows environment.
- Native read-only VBA write-surface audit completed for all three distinct H10 candidates. Each retained `ReadOnly=True`, macros/events disabled, `SOURCE HASH UNCHANGED: YES`, and `NO WORKBOOK WAS MODIFIED`.
- Each candidate still reports 1,569 static operation-pattern matches. Cross-candidate verifier: **GREEN**, three distinct workbooks, 44 nonempty component metadata records, 12 procedure-profile hashes and 9 candidate call-edge references. The pattern count is not a runtime write count.

**Call-edge detail (same in all three candidates):**
1. Five `ThisWorkbook` event calls: Open/Activate → `EnterApplicationShell`; Deactivate/BeforeClose → `ExitApplicationShell`; SheetActivate → `KeepApplicationShell`. These five calls were already matched at exact body-hash level to tracked `application_shell_vba.py::_EVENT_SPECS`.
2. `PATIENT_PLANNER.Worksheet_Change` has three branches, each referencing `CheckSingleAssignment` (native lines 16, 19, 22).
3. `modSchedulingConflict.ConfirmTherapistDoubleBooking` references `ConfigureConflict`.

The native component `Module11` declares `CheckSingleAssignment`, plus `BuildBusyMapsForCheck`, `GetPresentTherapists`, `GetTherapistLoad` and `GetSlotList`. Its 30 detected operation rows are worksheet references and cell-value/formula **accesses** only; no direct assignment or range/structure mutation candidate was detected in this component. Its literal worksheet references are `PATIENT_PLANNER`, `THERAPIST_ATTENDANCE` and `SETTINGS`. This supports classifying it as a read-side single-assignment/availability validation subsystem at the static-pattern level; it is not a complete control-flow proof.

`Module6` exposes `NormalizeTimeMS` and `NormalizePresenceMS`, with identical procedure fingerprints in the three candidates. No candidate VBA call-edge to either helper appeared in this report. This does **not** rule out formula/UDF references, external COM calls, or other invocation mechanisms not represented as a VBA identifier call.

**Next read-only refinement committed after this accepted bundle:**
- The auditor now searches for candidate references to *all* declared VBA procedures in every component with matched operation rows, rather than limiting those components to the two Module6 normalizers. This should reveal Module11's own helper call graph and other cross-module dependencies.
- The structural OOXML audit now counts worksheet-formula references to `NormalizeTimeMS` and `NormalizePresenceMS` without printing formula text, values, identities or cached values. Excel string literals are stripped before matching to avoid false positives.
- The one-command Windows gate now includes this formula scan on the same three frozen candidates and the focused regression for its new privacy-safe parser.

**Current status:** the just-reported 38/1,854 GREEN results validate the preceding revision, not the expanded call-graph/UDF-formula changes committed afterwards. Those newest changes require a fresh focused/full/native read-only gate. No frozen workbook has been modified or selected as a writable integration base.

Next product phase remains the whole-workbook demo. The real-patient source/review workbook and separate eligible pilot XLSM remain a distinct, operator-held admission gate. Do not merge real patient data into a synthetic fixture or treat the H10 candidates as writable targets.


## 2026-11-15 20:40 — expanded all-procedure call graph and formula-only UDF scan prepared

Following the accepted 20:31 bundle, two further **read-only** analysis improvements were committed:

1. `Python/tools/audit_vba_write_surfaces.py` now emits source-redacted `POSSIBLE_CALL_EDGE` candidates against **all declared VBA procedures** from every component with matched operation rows, not only the two Module6 normalizers. This is intended to expose Module11 helper calls and other cross-module edges. Edges are identifier references after quoted strings/comments are removed; they are not runtime execution proof. The synthetic tests now assert cross-procedure discovery and that the previous normalizer-only filter is absent.
2. `Python/tools/audit_integrated_workbook_structure.py` now counts formula references to `NormalizeTimeMS` and `NormalizePresenceMS` in the formula view, per sheet and per workbook, without printing formulas, values, cached results or identities. Quoted Excel string literals are removed before matching. Tests cover actual function calls and false-positive text constants.

`Scripts/run_h10_vba_call_edge_gate.ps1` was updated to include the new focused test file and run the read-only structure/UDF formula scan over the same three frozen H10 candidates. It requires 3/3 completed audits, unchanged source hashes and no-modification attestations, then adds the structure/UDF log to the single output bundle.

**Validation boundary:** these new commits are after the accepted 38-focused / 1,854-full result and have not yet been run by the operator. They must not inherit that earlier GREEN. The next local action is one execution of the updated PowerShell gate with Excel closed. The new script still does not save or modify any XLSM.

After this proof, proceed to the full synthetic integration workbook design/build. The real-source file named `DIOMIL_ΕΛΕΓΜΕΝΗ_ΜΕΤΑΦΟΡΑ_ΠΗΓΗΣ_ΠΡΟΣ_MASTER_20261112.xlsx` was not found among the currently available Project attachments; do not guess its local location or commit it. When the real-data admission milestone is reached, ask the operator to supply the reviewed source privately and follow `REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md`. The A:T column mapping already supplied in the conversation is retained as the operator's source-schema reference; the reclining-treatment replacement rule still requires the operator's fresh clinical instructions at the point of implementation.
