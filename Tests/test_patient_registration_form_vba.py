from rehab_excel.patient_registration_form_vba import (
    PATIENT_FORM_CODE,
    PATIENT_FORM_NAME,
)
from rehab_excel.registration_menu_vba import USERFORM_CODE


def test_patient_form_name_is_stable():
    assert PATIENT_FORM_NAME == "frmNewPatient"


def test_menu_patient_button_opens_patient_form():
    assert "Private Sub cmdPatient_Click()" in USERFORM_CODE
    assert "frmNewPatient.Show" in USERFORM_CODE


def test_patient_form_contains_required_validation():
    assert "Το Patient ID είναι υποχρεωτικό." in PATIENT_FORM_CODE
    assert "Το ονοματεπώνυμο είναι υποχρεωτικό." in PATIENT_FORM_CODE
    assert "Private Function ValidateForm() As Boolean" in PATIENT_FORM_CODE


def test_patient_form_reads_room_and_status_from_settings():
    assert 'ThisWorkbook.Worksheets("SETTINGS")' in PATIENT_FORM_CODE
    assert "Cells(ws.Rows.Count, 8)" in PATIENT_FORM_CODE
    assert "Cells(rowIndex, 8)" in PATIENT_FORM_CODE
    assert "Cells(ws.Rows.Count, 6)" in PATIENT_FORM_CODE
    assert "Cells(rowIndex, 6)" in PATIENT_FORM_CODE


def test_patient_form_exposes_all_contract_fields():
    for control_name in (
        "txtPatientID",
        "txtDisplayName",
        "cboRoom",
        "chkInfectious",
        "cboStatus",
    ):
        assert control_name in PATIENT_FORM_CODE


def test_patient_form_does_not_write_workbook_yet():
    assert ".Save" not in PATIENT_FORM_CODE
    assert "Cells(" not in PATIENT_FORM_CODE.replace("Cells(ws.Rows.Count", "").replace("Cells(rowIndex", "")
    assert "registration backend" in PATIENT_FORM_CODE
