from rehab_core.models import PatientType
from rehab_core.registration import NewPatientRequest, validate_new_patient
from rehab_excel.patient_registration import (
    HOSPITAL_MRN_HEADERS,
    OUTPATIENT_SCHEDULE_HEADERS,
    PATIENT_TYPE_HEADERS,
)


def test_legacy_new_patient_request_defaults_to_inpatient():
    request = NewPatientRequest("P1", "ΑΣΘΕΝΗΣ")

    assert request.patient_type == PatientType.INPATIENT
    assert request.hospital_mrn is None


def test_outpatient_request_accepts_optional_hospital_mrn():
    request = NewPatientRequest(
        "O1",
        "ΕΞΩΤΕΡΙΚΟΣ",
        patient_type=PatientType.OUTPATIENT,
        hospital_mrn="MRN-77",
    )

    check = validate_new_patient(request)

    assert check.allowed is True
    assert request.hospital_mrn == "MRN-77"


def test_outpatient_request_rejects_inpatient_room_and_infectious_flag():
    request = NewPatientRequest(
        "O1",
        "ΕΞΩΤΕΡΙΚΟΣ",
        room="101",
        infectious=True,
        patient_type=PatientType.OUTPATIENT,
    )

    check = validate_new_patient(request)

    assert check.allowed is False
    assert {issue.code for issue in check.issues} == {
        "outpatient_infectious_not_allowed",
        "outpatient_room_not_allowed",
    }


def test_patient_registration_schema_headers_match_reader_contract():
    assert "PatientType" in PATIENT_TYPE_HEADERS
    assert "HospitalMRN" in HOSPITAL_MRN_HEADERS
    assert OUTPATIENT_SCHEDULE_HEADERS == (
        "PatientID",
        "Ασθενής",
        "Θεραπεία",
        "Ώρα",
        "Ημέρες",
        "Θεραπευτής",
    )
