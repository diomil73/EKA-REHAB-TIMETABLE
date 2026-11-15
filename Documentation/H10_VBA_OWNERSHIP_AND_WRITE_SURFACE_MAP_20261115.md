# DIOMIL — VBA ownership and write-surface reconciliation

Date: 2026-11-15
Branch: work/afternoon-medical-scope-20261115
Status: native inventory and sanitized write-surface audit complete for three distinct candidates; source ownership mapping partial; four native-only components remain for source-level behavior review.

## 1. Native inventory evidence

Operator report: DIOMIL_H10_VBA_INVENTORY_FINAL_20261115.txt.

The report completed all distinct candidate inventories:
- UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm: 67 components.
- UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm: 67 components.
- UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm: 66 components.
- UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm: byte-identical duplicate of the preceding CC probe; duplicate skipped with original audit status 0.

The first two inventories have the same 67 component names. The third has 66; the report's one differing document-module name is corrupted by the Windows output code page and must be resolved by the Unicode-safe write-surface audit. The fourth file is not a fourth distinct VBA project.

For the 67-component recurring-ID candidate:
- 27 DOCUMENT_MODULE components.
- 30 STANDARD_MODULE components.
- 10 USERFORM components.
- 12,185 total reported code lines across components.
- 30 named application components can be mapped to tracked Python source/generators (9,150 code lines).
- 37 components remain unresolved by source ownership (3,035 code lines): 10 legacy/generic standard modules and 27 worksheet/document modules. Do not assume they are obsolete.

Native inspection controls:
- ReadOnly=True.
- Macros disabled before workbook open.
- Excel events disabled.
- VBA project protection state reported as 0 (not protected).
- Source SHA-256 unchanged for every distinct candidate.
- No workbook was modified.
- Manual Calculation mode was unavailable in this Excel COM environment; this was recorded as a warning, not silently treated as successful.

The uploaded final report is the native evidence. The focused-test and full-regression logs from the same gated operator run have not been separately attached here, so their exact pass counts are not entered in this register. The operator script only reaches the native inventory stage after both test gates return zero.

## 2. Tracked component ownership map

The following mappings are confirmed by matching the embedded component names to constants/generators in the current branch. They are source-ownership links, not yet proof that every embedded line in every candidate equals the current generator output.

| Embedded VBA component(s) | Tracked generator / installer | Relevant regression coverage |
|---|---|---|
| modRegistrationMenu, frmRegistrationMenu | Python/rehab_excel/registration_menu_vba.py | Tests/test_registration_menu_vba.py |
| modApplicationShell | Python/rehab_excel/application_shell_vba.py | Tests/test_application_shell_vba.py; Tests/test_reinstall_application_shell.py |
| modMasterToolbar | Python/rehab_excel/master_toolbar_vba.py | Tests/test_master_toolbar_vba.py; Tests/test_master_toolbar_pinned_view.py |
| modDailyInputAction | Python/rehab_excel/daily_input_action_vba.py | Tests/test_daily_input_action_vba.py |
| modGlobalReplacementAction | Python/rehab_excel/global_replacement_action_vba.py | Tests/test_global_replacement_action_vba.py |
| modDoubleBookingReview, frmDoubleBookingReview | Python/rehab_excel/double_booking_review_vba.py | Tests/test_double_booking_review_vba.py; Tests/test_install_double_booking_review_ui.py |
| frmNewPatient | Python/rehab_excel/patient_registration_form_vba.py | Tests/test_patient_registration_form_vba.py |
| frmNewTherapist | Python/rehab_excel/therapist_registration_form_vba.py | Tests/test_therapist_registration_form_vba.py |
| frmNewStudent | Python/rehab_excel/student_registration_form_vba.py | Tests/test_student_registration_form_vba.py |
| modCloseDayAction | Python/rehab_excel/close_day_action_vba.py | Tests/test_close_day_action_vba.py |
| modTherapistWorkloadStrip | Python/rehab_excel/therapist_workload_strip_vba.py | Tests/test_therapist_workload_strip_vba.py |
| modReplacementOperator, frmReplacementDecision | Python/rehab_excel/replacement_operator_vba.py | Tests/test_replacement_operator_recovery.py |
| modDailyChangeLedger | Python/rehab_excel/daily_change_ledger_in_place_vba.py; shared ledger contract in Python/rehab_excel/daily_change_ledger_vba.py | Tests/test_daily_change_ledger_in_place_vba.py; Tests/test_daily_change_ledger_vba.py |
| modPatientRegistrationBridge | Python/rehab_excel/patient_registration_bridge_vba.py | Tests/test_patient_registration_bridge_vba.py |
| frmEditPatient | Python/rehab_excel/patient_edit_form_vba.py and Python/rehab_excel/patient_edit_flow_vba.py | Tests/test_patient_edit_flow_vba.py |
| modOutpatientSchedule, frmOutpatientSchedule | Python/rehab_excel/outpatient_schedule_form_vba.py | Tests/test_outpatient_schedule_form_vba.py |
| modOutpatientScheduleEdit | Python/rehab_excel/outpatient_schedule_edit_vba.py | Tests/test_outpatient_schedule_edit_vba.py |
| modOutpatientScheduleManager, frmOutpatientScheduleManager | Python/rehab_excel/outpatient_schedule_manager_vba.py | Tests/test_outpatient_schedule_manager_vba.py |
| modSchedulingConflict, frmSchedulingConflict | Python/rehab_excel/scheduling_conflict_popup_vba.py | Tests/test_scheduling_conflict_popup_vba.py |
| modOutpatientInSession | Python/rehab_excel/outpatient_in_session_vba.py | Tests/test_outpatient_in_session_vba.py |
| modDayAmendmentAction | Python/rehab_excel/day_amendment_action_vba.py | Tests/test_day_amendment_action_vba.py |
| modUnifiedDaySave | Python/tools/install_unified_day_save_ui.py | Tests/test_unified_day_save_ui.py |
| modPatientHistoryReport | Python/rehab_excel/patient_history_report_vba.py | Tests/test_patient_history_reports_h7.py |
| modTherapistHistoryReport | Python/rehab_excel/therapist_history_report_vba.py | Tests/test_therapist_history_reports.py |

## 3. Unresolved components: mandatory reconciliation

The remaining 37 components are not yet assigned to tracked source:
- 10 STANDARD_MODULE components with generic names Module1, Module2, Module3, Module4, Module5, Module6, Module7, Module8, Module10 and Module11.
- 27 DOCUMENT_MODULE components corresponding to workbook/sheet code-behind. Their names were corrupted in the captured console output, so no worksheet identity is inferred from that output.

The 66-component CC project lacks one zero-code-line document module compared with the 67-component candidates. This may be a sheet code-behind component, but its identity and operational significance are not yet established.

These components may contain legacy logic, workbook events, buttons, protection behavior, formulas/projection hooks or other critical paths. They must be audited, not deleted or dismissed.

## 4. Write-surface audit hardening

Python/tools/audit_vba_write_surfaces.py was hardened to:
- launch an isolated Excel process;
- disable macros (AutomationSecurity=3), events and screen updates before opening;
- open workbooks with ReadOnly=True, UpdateLinks=0 and AddToMru=False;
- inspect component procedures and worksheet write/protection/visibility operations;
- output only component/procedure/line/operation metadata and target worksheet names, not raw VBA source lines;
- escape Unicode component and worksheet names so Greek names survive PowerShell output encoding;
- hash the workbook before and after and fail if the fingerprint changes;
- close with SaveChanges=False;
- accept multiple workbook paths and skip byte-identical duplicates while preserving any prior nonzero status.

This is a read-only source inspection. It does not run macros, change Trust Center settings, edit or save a workbook.

## 5. Next operator gate

After pulling this branch, run the focused tests, full source regression, and then the sanitized write-surface audit against the three distinct native candidates:
1. UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm
2. UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm
3. UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm

The byte-identical UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm need not be audited a second time.

The write-surface report must end with VBA WRITE-SURFACE AUDIT COMPLETE and SOURCE HASH UNCHANGED: YES for all three distinct candidates. Any COM error, incomplete output, hash change or missing workbook is a hard stop. Do not enable Trust Center access or unlock any project to force completion.

## 6. After the operator gate

1. Reconcile the 37 unresolved components against worksheet CodeNames, worksheet names, procedures, tracked VBA generators and installer history.
2. Build a sheet-by-sheet write matrix: data writes, formula writes, clear/delete/insert, protect/unprotect, visibility, navigation, Save/SaveAs and event procedures.
3. Reconcile that matrix with Documentation/WORKSHEET_PROTECTION_WRITE_PATH_MAP_2026-11-04.md. Do not implement blanket protection.
4. Only then select the trusted synthetic baseline and design a fresh, uniquely named whole-workbook synthetic integration candidate.
5. Keep all four H10 source candidates frozen and unchanged.

## 7. Demo acceptance boundary

The end goal remains a functional synthetic demo with the agreed patient/treatment source mapping, stable patient identities, real therapist/treatment/day/time relationships, and integrated daily operations. The actual patient source and identity mapping remain a separate private-data admission gate. This VBA inventory does not certify the workbook as a whole-program synthetic fixture and does not authorize patient-data import.


## First write-surface audit result — 2026-11-15

The uploaded 18:50 audit was safely incomplete on all three distinct candidates. The same module, modTherapistHistoryReport, contained 265 matched source lines, while the earlier auditor emitted only 200 per component. The reports printed 1,331 operations and marked output incomplete; component headers indicated 1,396 total matched operations. All three workbooks were opened read-only with macros/events disabled, their SHA-256 fingerprints remained unchanged, and no workbook was modified.

The auditor cap has since been raised to 500 per component and 5,000 total. A regression test protects the known 265-per-component and 1,396-total inventory. The classifier now distinguishes direct Value/Value2/Formula assignment from read/access lines. This remains static source-pattern evidence and requires review before treating any candidate as a confirmed mutation.

Rerun the focused tests, full regression, and all three distinct native candidates with the updated branch. Accept only complete, uncapped reports with unchanged source hashes.


## Completed uncapped native audit — 2026-11-15

The corrected auditor completed all three distinct candidates read-only with macros/events disabled and unchanged SHA-256 fingerprints. Each had 1,396 matched source lines and no output truncation. Component counts were 67, 67, and 66 respectively.

The audit identified 35 components with at least one candidate source-pattern match per workbook. Inventory reconciliation normalizes two worksheet-module name encoding aliases (component 008 / worksheet 7 and component 061 / worksheet 22). This leaves 32 components with no matched operation pattern: 23 have zero code lines and 9 have non-empty code but no current detector match. No-match does not prove absence of side effects; the 9 non-empty components need targeted source review.

Per workbook, candidate operation counts are: worksheet_reference 80; cell_value_or_formula_access 373; cell_value_or_formula_write 282; range_or_structure_mutation 111; format_or_other_access_or_write 460; visibility_navigation_or_action_binding 30; workbook_save 26; protection_unprotect 17; protection_protect 17. Sum: 1,396. These are static candidate line classifications, not confirmed runtime writes. Review direct assignments, range/structure changes, save and protection calls before drawing ownership conclusions.

Full regression: 1,824 passed, 1 skipped (symlinks unavailable on Windows), 26 warnings.


### Reconciliation count clarification

The 32-component no-match reconciliation (23 zero-code + 9 non-empty) is grounded in the 67-component inventory candidate and applies to the two 67-component reports. The CC read-only probe has 66 components and the same 35 matched components, so its arithmetic remainder is 31, not 32. Its candidate-specific component inventory has not yet been reconciled item-by-item; do not project the 23/9 breakdown onto that 66-component workbook. The two worksheet-module display-name aliases are normalized only when matching against the 67-component inventory.


## Targeted source review of components without worksheet-write matches

Tracked-source review establishes the following:
- modRegistrationMenu: entry macro displays frmRegistrationMenu; form handlers route to patient, therapist, student and outpatient UI. No direct worksheet writes in the module.
- modOutpatientSchedule: entry macro displays frmOutpatientSchedule. The form reads PATIENTS and SETTINGS, then submits a request to outpatient_schedule_bridge_cli.py. A preview and source SHA-256 are returned and passed to authoritative_commit_worker_cli.py. Actual workbook writeback is delegated to the Python preview/commit path.
- modOutpatientScheduleManager: entry macro displays the manager form. The form reads OUTPATIENT_SCHEDULE and delegates create/edit actions to the outpatient form/edit modules.
- modSchedulingConflict: modal therapist double-booking confirmation. Safe default is change/cancel; it returns an explicit Accepted decision. No direct workbook writes.
- modDoubleBookingReview: performs external file I/O (temporary JSON/TSV creation/deletion), launches Python prepare/apply/authoritative-commit tools, and coordinates closing/reopening the workbook around authoritative commit. It is not a no-op; the previous worksheet-only scanner did not detect these external side effects.

Related regression contracts: Tests/test_registration_menu_vba.py; Tests/test_double_booking_review_vba.py; Tests/test_outpatient_schedule_form_vba.py; Tests/test_outpatient_schedule_manager_vba.py; Tests/test_scheduling_conflict_popup_vba.py. These prove source-level contracts, not native end-to-end runtime acceptance.

Still unresolved: generic Module5 (1 line), generic Module6 (26 lines), and two worksheet/document modules (19 and 12 lines). No tracked generator source has been found for them. A new sanitized auditor revision reports procedure identifiers and worksheet names for non-empty no-match components and detects external file I/O, process launch and workbook-close surfaces. Run the updated audit before assigning ownership.


## Latest continuation gate — 2026-11-15

The VBA auditor was further extended after the first targeted source review:
- emits NO_PATTERN_MATCH_COMPONENT records with escaped component name, type, line count, worksheet name when applicable, and procedure identifiers only; it never prints source text;
- detects external file create/write/delete/read, process launches, and ThisWorkbook.Close in addition to worksheet and workbook write surfaces;
- excludes procedure declaration lines from operation matches;
- recognizes .Run(...) calls with or without whitespace.

New focused regression tests cover procedure inventory, external I/O/process/lifecycle classification, and declaration-line exclusion. These latest changes have not yet been executed on the operator's Windows/Excel machine. The previous 1,824-pass full regression and completed native audit predate these latest changes and must not be presented as validation of them.

Next operator gate: pull branch work/afternoon-medical-scope-20261115; run Tests/test_audit_vba_write_surfaces.py and the existing integrated-structure/VBA-inventory tests; run full pytest; then rerun the sanitized audit on the same three distinct H10 candidates. Upload the focused/full regression logs and final write-surface report. The new report should expose procedure identifiers for the remaining unmatched non-empty components and separately classify the external file/process operations in modDoubleBookingReview. Preserve ReadOnly=True, disabled macros/events, unchanged source hashes and no workbook modifications.


## Operator validation received — 2026-11-15 19:34 UTC

Evidence bundle: DIOMIL_VBA_SOURCE_REVIEW_20261115_193409.txt

Results:
- Focused auditor/inventory/structure tests: 24 passed.
- Full pytest: 1,829 passed, 1 skipped (symlinks unavailable), 26 warnings, 29.64 seconds.
- Native read-only VBA write-surface audit completed for all three H10 candidates: UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm, UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm, and UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm.
- Each audit reports 1,569 matched operations, SOURCE HASH UNCHANGED: YES, and NO WORKBOOK WAS MODIFIED.
- Read-only safeguards reported: ReadOnly=True, macros disabled before open, Excel events disabled.

The same eight no-pattern components appear in each candidate: ThisWorkbook document module (19 lines; Workbook_Open/Activate/Deactivate/SheetActivate/BeforeClose), worksheet module for PATIENT_PLANNER (25 lines; Worksheet_Change), Module6 (26 lines; NormalizeTimeMS and NormalizePresenceMS), Module5 (1 line; no procedure), modRegistrationMenu, modOutpatientSchedule, modOutpatientScheduleManager, and modSchedulingConflict. The four named UI/navigation components were source-reviewed separately and are not direct worksheet-write procedures. The workbook event module, PATIENT_PLANNER Worksheet_Change, and Module6 still need explicit ownership/behavior review; Module5 should be checked for whitespace/attribute-only/empty content before classifying as inert.

The modDoubleBookingReview external file/process/workbook-close behavior is now visible in the audit, confirming why the former worksheet-only scan was incomplete. Do not interpret the 1,569 pattern matches as 1,569 confirmed writes: the report includes accesses, reads, formatting/navigation, and writes. The classification is a static surface inventory, not runtime proof.

Next: review the three remaining behavior-bearing no-pattern components (ThisWorkbook events, PATIENT_PLANNER Worksheet_Change, Module6 time/presence normalization); classify Module5; then consolidate component ownership, side-effect type, target sheets/files/processes, and test/native evidence into the master write-surface map. No further workbook mutation or native write trial is authorized by this audit.


## Native VBA audit and regression results — 2026-11-15 19:34 UTC

Operator uploaded DIOMIL_VBA_SOURCE_REVIEW_20261115_193409.txt. Evidence reviewed:
- Focused tests: 24 passed in 2.50s.
- Full pytest: 1,829 passed, 1 skipped, 26 warnings in 29.64s (1,830 collected).
- Three distinct VBA candidates completed the read-only write-surface audit. Each reports ReadOnly=True, macros disabled before open, Excel events disabled, SOURCE HASH UNCHANGED: YES, and NO WORKBOOK WAS MODIFIED.
- Candidate hashes: RECURRING_IDS = 4df997b6a5e5b0b56b37127ee00dd8710cc26c7c118f1742d58c9969bcbe7afa; NATIVE_WRITE = fa9765da7079137bcb1ff935c3d08bc84b409c8335f9cd719f3583b89e4fc67d; CC_NATIVE_READONLY_PROBE = 1bd5e58e905f91acb55e5a706cd8b4563b5a98c170a9ec60fe76c95617420b13. Candidate counts are 67, 67 and 66 VBA components respectively.
- The separate native inventory reports Excel calculation mode could not be set to Manual because of a COM error. This is a warning, not a failed audit; retain it as an explicit limitation. ReadOnly and macro/event safeguards were confirmed.
- The write-surface audit detects 1,569 matched operations in each candidate. Treat this as a static pattern inventory, not 1,569 proven mutations or a runtime execution trace.

Unmatched-component investigation now has useful identities:
- The 19-line DOCUMENT_MODULE is the workbook event module: Workbook_Open, Workbook_Activate, Workbook_Deactivate, Workbook_SheetActivate, Workbook_BeforeClose.
- The 25-line worksheet document module maps to PATIENT_PLANNER and contains Worksheet_Change.
- Module6 (26 lines) contains NormalizeTimeMS and NormalizePresenceMS. These are helper procedures with no current write-surface pattern match; inspect their actual implementation/consumers before assigning a stronger classification.
- Module5 contains one code line and no recognized procedure.
- modRegistrationMenu, modOutpatientSchedule, modOutpatientScheduleManager and modSchedulingConflict are now identified by entry procedure names and their source-level roles are documented in the targeted review.
- modDoubleBookingReview is positively identified as an external side-effect orchestrator: temporary file create/delete/read/write, Python process launches and workbook close/reopen coordination. The scanner reports these explicitly (16 matches in the module), so it must be treated as a write/lifecycle owner even though it has no direct worksheet assignment in the entry module.

Important audit limitation: procedure names and pattern matches identify ownership candidates, not full semantics. The two event procedures, Module6 helpers and one-line Module5 still require implementation-level review and caller/callee tracing. Do not mark the VBA ownership map fully closed yet. The inventory emits component names in escaped Unicode where applicable; no source text or cell values were printed.

Next step: complete the remaining source-level trace for Workbook events, PATIENT_PLANNER.Worksheet_Change, Module6 and Module5 using tracked VBA generator/source where available, and record unresolved native-only code honestly. Do not modify source workbooks. The regression and read-only audit gate is green; proceed with documentation and static ownership analysis, not production workbook changes.


## Current reconciliation — supersedes earlier no-pattern counts

Evidence bundle: DIOMIL_VBA_SOURCE_REVIEW_20261115_193409.txt.

Three distinct candidates completed the latest expanded native audit. Each produced 1,569 static operation-pattern matches, with ReadOnly=True, macros disabled, Excel events disabled, unchanged SHA-256, and no workbook modification. Full regression: 1,829 passed, 1 skipped, 26 warnings; focused gate: 24 passed. Calculation mode could not be switched to Manual due to a COM error; retain as a limitation.

Do not conflate the two different counts:
- **37 components unresolved by tracked-source ownership** in the 67-component inventory remains a generator/source-mapping gap (10 generic standard modules and 27 document modules). This is not the count of components with no operation-pattern match.
- **8 components with no operation-pattern match** in each latest candidate are: ThisWorkbook, PATIENT_PLANNER worksheet module, Module6, Module5, modRegistrationMenu, modOutpatientSchedule, modOutpatientScheduleManager, and modSchedulingConflict. Four named UI/navigation modules have tracked source and have been reviewed. Four native-only components still require behavior review: ThisWorkbook events, PATIENT_PLANNER.Worksheet_Change, Module6 helper implementations/consumers, and Module5's one line.

The expanded detector now finds modDoubleBookingReview external file I/O, Python process launches and workbook close lifecycle. It must be included in the write/lifecycle ownership map even though its entry module does not directly write worksheet cells.

The 1,569 matches per workbook are static source-pattern locations, not confirmed runtime writes or proof of call reachability. Native inventory/audit is complete; semantic behavior review and full synthetic integration acceptance are not.

Next: inspect the four native-only components through a local, read-only, source-redacting diagnostic that reports only procedure identifiers, statement-category counts and per-procedure hashes; compare across all three candidates. No source workbook edits or saves.
