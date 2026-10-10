from openpyxl import Workbook

from rehab_core.models import Patient
from rehab_core.registration import NewPatientRequest
from rehab_excel.patient_registry_source import read_patient_registry
from rehab_excel.patient_registration import RESPONSIBLE_DOCTOR_HEADERS


def _save(tmp_path, headers, row):
    path = tmp_path / "patients.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "PATIENTS"
    ws.append(headers)
    ws.append(row)
    wb.save(path)
    return path


def test_patient_model_and_registration_request_keep_doctor_optional():
    patient = Patient("P1", "ΑΣΘΕΝΗΣ")
    request = NewPatientRequest("P2", "ΝΕΟΣ")

    assert patient.responsible_doctor is None
    assert request.responsible_doctor is None


def test_legacy_registry_without_doctor_stays_backward_compatible(tmp_path):
    path = _save(
        tmp_path,
        ["PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση"],
        ["P1", "A01", "ΑΣΘΕΝΗΣ", "Ο", "Ενεργός"],
    )

    [patient] = read_patient_registry(path)
    assert patient.responsible_doctor is None


def test_registry_reads_responsible_doctor_from_optional_header(tmp_path):
    path = _save(
        tmp_path,
        [
            "PatientID",
            "Θάλαμος",
            "Ασθενής",
            "Λοιμώδης",
            "Κατάσταση",
            "ResponsibleDoctor",
        ],
        ["P1", "A01", "ΑΣΘΕΝΗΣ", "Ο", "Ενεργός", "Νικολάου"],
    )

    [patient] = read_patient_registry(path)
    assert patient.responsible_doctor == "Νικολάου"


def test_registration_supports_greek_and_english_doctor_headers():
    assert "ResponsibleDoctor" in RESPONSIBLE_DOCTOR_HEADERS
    assert "Υπεύθυνος Ιατρός" in RESPONSIBLE_DOCTOR_HEADERS
    assert "Υπεύθυνος Γιατρός" in RESPONSIBLE_DOCTOR_HEADERS
