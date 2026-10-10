from rehab_excel.patient_registration_bridge_vba import (
    BRIDGE_MODULE_CODE,
    BRIDGE_MODULE_NAME,
    FORM_BRIDGE_CODE,
    PATIENT_FORM_NAME,
)
from rehab_excel.patient_registration_bridge_diagnostics import (
    _patched_form_code,
    _patched_module_code,
)


def test_bridge_targets_existing_patient_form():
    assert PATIENT_FORM_NAME == "frmNewPatient"
    assert "Private Sub cmdSave_Click()" in FORM_BRIDGE_CODE


def test_bridge_uses_in_process_registration_not_detached_worker():
    assert "RegisterPatientInWorkbook" in FORM_BRIDGE_CODE
    assert "registration_transaction_worker_cli.py" not in FORM_BRIDGE_CODE
    assert "WScript.Shell" not in FORM_BRIDGE_CODE
    assert "ThisWorkbook.Close" not in FORM_BRIDGE_CODE


def test_bridge_passes_patient_values_directly_to_workbook_registration():
    assert "txtHospitalMRN.Text" in FORM_BRIDGE_CODE
    assert "txtDisplayName.Text" in FORM_BRIDGE_CODE
    assert "cboRoom.Value" in FORM_BRIDGE_CODE
    assert "chkInfectious.Value" in FORM_BRIDGE_CODE
    assert "cboStatus.Value" in FORM_BRIDGE_CODE


def test_bridge_requires_room_only_for_inpatients():
    assert 'If patientType <> "Εξωτερικός" Then' in FORM_BRIDGE_CODE
    assert 'roomValue = ""' in FORM_BRIDGE_CODE
    assert "Ο θάλαμος είναι υποχρεωτικός για εσωτερικό ασθενή." in FORM_BRIDGE_CODE


def test_fast_path_updates_registry_planner_master_and_saves_once():
    assert "Public Function RegisterPatientInWorkbook" in BRIDGE_MODULE_CODE
    assert 'ThisWorkbook.Worksheets("PATIENTS")' in BRIDGE_MODULE_CODE
    assert 'ThisWorkbook.Worksheets("PATIENT_PLANNER")' in BRIDGE_MODULE_CODE
    assert "RefreshPlannerRow" in BRIDGE_MODULE_CODE
    assert "RebuildMasterFast" in BRIDGE_MODULE_CODE
    assert "ThisWorkbook.Save" in BRIDGE_MODULE_CODE
    assert "Workbooks.Open" not in BRIDGE_MODULE_CODE
    assert "Application.Quit" not in BRIDGE_MODULE_CODE


def test_fast_path_keeps_outpatients_out_of_master_projection():
    assert "If Not IsOutpatient(patientType) Then" in BRIDGE_MODULE_CODE
    assert "RebuildMasterFast patients, typeCol, doctorCol" in BRIDGE_MODULE_CODE


def test_fast_path_preserves_manual_calculation_guard_and_restores_excel_state():
    assert "previousCalculation = Application.Calculation" in BRIDGE_MODULE_CODE
    assert "Application.Calculation = xlCalculationManual" in BRIDGE_MODULE_CODE
    assert "Application.Calculation = previousCalculation" in BRIDGE_MODULE_CODE
    assert "Application.EnableEvents = previousEvents" in BRIDGE_MODULE_CODE
    assert "Application.ScreenUpdating = previousScreenUpdating" in BRIDGE_MODULE_CODE


def test_diagnostic_bridge_reports_exact_failure_stage():
    code = _patched_module_code()
    assert 'Dim stage As String' in code
    assert 'stage = "write PATIENTS row " & CStr(targetRow)' in code
    assert 'stage = "refresh PATIENT_PLANNER row " & CStr(targetRow)' in code
    assert 'stage = "rebuild MASTER"' in code
    assert 'stage = "save workbook"' in code
    assert 'errorText = "Stage: " & stage & vbCrLf & Err.Description' in code


def test_diagnostic_bridge_uses_logical_registry_row_not_blank_table_tail():
    code = _patched_module_code()
    assert "Private Function LastLogicalPatientRow" in code
    assert "Value2" in code
    assert "LastLogicalPatientRow(ws, 1)" in code
    assert "LastLogicalPatientRow(ws, 3)" in code


def test_successful_registration_returns_to_master_without_closing_workbook():
    code = _patched_form_code()
    assert "Unload frmRegistrationMenu" in code
    assert "Unload Me" in code
    assert "GoToMaster" in code
    assert "ThisWorkbook.Close" not in code
    assert "Application.Quit" not in code


def test_bridge_keeps_standard_module_name_for_compatibility():
    assert BRIDGE_MODULE_NAME == "modPatientRegistrationBridge"
