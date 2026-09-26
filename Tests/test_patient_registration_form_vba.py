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


def test_patient_form_requires_type_and_name_not_manual_patient_id():
    assert "Ο τύπος ασθενή είναι υποχρεωτικός." in PATIENT_FORM_CODE
    assert "Το ονοματεπώνυμο είναι υποχρεωτικό." in PATIENT_FORM_CODE
    assert "Το Patient ID είναι υποχρεωτικό." not in PATIENT_FORM_CODE
    assert 'txtPatientID.Text = "Αυτόματο κατά την αποθήκευση"' in PATIENT_FORM_CODE
    assert "txtPatientID.Locked = True" in PATIENT_FORM_CODE
    assert "Private Function ValidateForm() As Boolean" in PATIENT_FORM_CODE


def test_patient_form_supports_inpatient_and_outpatient_types():
    assert 'cboPatientType.AddItem "Εσωτερικός"' in PATIENT_FORM_CODE
    assert 'cboPatientType.AddItem "Εξωτερικός"' in PATIENT_FORM_CODE
    assert "Private Sub ApplyPatientTypeRules()" in PATIENT_FORM_CODE
    assert 'cboPatientType.Value <> "Εξωτερικός"' in PATIENT_FORM_CODE


def test_patient_form_exposes_optional_hospital_mrn():
    assert "txtHospitalMRN" in PATIENT_FORM_CODE
    assert "ΑΜ Νοσοκομείου" in PATIENT_FORM_CODE
    assert "μπορεί να συμπληρωθεί και αργότερα" in PATIENT_FORM_CODE


def test_patient_form_reads_room_and_status_from_settings():
    assert 'ThisWorkbook.Worksheets("SETTINGS")' in PATIENT_FORM_CODE
    assert "Cells(ws.Rows.Count, 8)" in PATIENT_FORM_CODE
    assert "Cells(rowIndex, 8)" in PATIENT_FORM_CODE
    assert "Cells(ws.Rows.Count, 6)" in PATIENT_FORM_CODE
    assert "Cells(rowIndex, 6)" in PATIENT_FORM_CODE


def test_outpatient_disables_inpatient_only_fields():
    for control_name in (
        "cboRoom.Enabled = isInpatient",
        "chkInfectious.Enabled = isInpatient",
        "cboStatus.Enabled = isInpatient",
    ):
        assert control_name in PATIENT_FORM_CODE
    assert "δεν εμφανίζεται στα φύλλα νοσηλευομένων" in PATIENT_FORM_CODE


def test_patient_form_does_not_write_workbook_yet():
    assert ".Save" not in PATIENT_FORM_CODE
    assert "registration backend" in PATIENT_FORM_CODE
