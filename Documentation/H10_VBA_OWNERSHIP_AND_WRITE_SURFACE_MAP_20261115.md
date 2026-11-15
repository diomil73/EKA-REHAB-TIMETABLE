# DIOMIL — VBA ownership and write-surface reconciliation

Date: 2026-11-15
Branch: work/afternoon-medical-scope-20261115
Status: native inventory complete; ownership mapping preliminary; write-surface operator audit pending.

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
