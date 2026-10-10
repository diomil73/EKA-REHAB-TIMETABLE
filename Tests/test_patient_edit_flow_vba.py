from rehab_excel.patient_edit_flow_vba import _menu_code_with_patient_edit
from rehab_excel.patient_edit_form_vba import PATIENT_EDIT_FORM_CODE, PATIENT_EDIT_FORM_NAME
from rehab_excel.patient_registration_vertical_slice import _vertical_module_code


def test_patient_edit_form_is_stable_and_loads_existing_patients():
    assert PATIENT_EDIT_FORM_NAME == "frmEditPatient"
    assert "Private Sub LoadPatients()" in PATIENT_EDIT_FORM_CODE
    assert "Private Sub cboPatient_Change()" in PATIENT_EDIT_FORM_CODE
    assert "LoadPatientById patientId" in PATIENT_EDIT_FORM_CODE
    assert 'txtPatientID.Locked = True' in PATIENT_EDIT_FORM_CODE


def test_patient_edit_form_supports_registry_field_corrections():
    assert "txtDisplayName" in PATIENT_EDIT_FORM_CODE
    assert "txtHospitalMRN" in PATIENT_EDIT_FORM_CODE
    assert "txtResponsibleDoctor" in PATIENT_EDIT_FORM_CODE
    assert "cboRoom" in PATIENT_EDIT_FORM_CODE
    assert "chkInfectious" in PATIENT_EDIT_FORM_CODE
    assert "cboStatus" in PATIENT_EDIT_FORM_CODE
    assert "UpdatePatientInWorkbook" in PATIENT_EDIT_FORM_CODE


def test_patient_edit_form_keeps_outpatient_doctor_room_and_infectious_disabled():
    assert "lblRoom.Enabled = Not isOutpatient" in PATIENT_EDIT_FORM_CODE
    assert "txtResponsibleDoctor.Enabled = Not isOutpatient" in PATIENT_EDIT_FORM_CODE
    assert 'txtResponsibleDoctor.Text = ""' in PATIENT_EDIT_FORM_CODE


def test_registry_menu_exposes_patient_edit_action():
    code = _menu_code_with_patient_edit()
    assert 'StyleMenuButton cmdEditPatient, "Επεξεργασία ασθενή", 113' in code
    assert "Private Sub cmdEditPatient_Click()" in code
    assert "frmEditPatient.Show" in code


def test_in_session_edit_backend_updates_same_patient_id_and_rebuilds_master():
    code = _vertical_module_code()
    assert "Public Sub UpdatePatientInWorkbook" in code
    assert "FindPatientRowById(patients, patientId)" in code
    assert "patients.Cells(targetRow, 3).Value = Trim$(displayName)" in code
    assert "patients.Cells(targetRow, doctorCol).Value = Trim$(responsibleDoctor)" in code
    assert "RefreshPlannerRow planner, targetRow" in code
    assert "RebuildMasterFast patients, typeCol, doctorCol" in code
    assert "ThisWorkbook.Save" in code
    assert "ThisWorkbook.Close" not in code
    assert "Application.Quit" not in code


def test_edit_backend_rejects_duplicate_hospital_mrn_except_same_record():
    code = _vertical_module_code()
    assert "If rowIndex <> targetRow Then" in code
    assert "Ο ΑΜ Νοσοκομείου υπάρχει ήδη σε άλλον ασθενή." in code
