from rehab_excel.patient_registration_bridge_vba import (
    BRIDGE_MODULE_CODE,
    BRIDGE_MODULE_NAME,
    FORM_BRIDGE_CODE,
    PATIENT_FORM_NAME,
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


def test_bridge_keeps_standard_module_name_for_compatibility():
    assert BRIDGE_MODULE_NAME == "modPatientRegistrationBridge"
