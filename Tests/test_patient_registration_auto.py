from rehab_core.models import Patient
from rehab_core.registration import NewPatientRequest
from rehab_excel.patient_registration_auto import resolve_patient_id_request


def test_blank_patient_id_is_allocated_sequentially():
    request = NewPatientRequest(patient_id="", display_name="ΝΕΟΣ")
    resolved = resolve_patient_id_request(
        request,
        existing_patients=[Patient("94", "A"), Patient("TEST-X", "B")],
    )
    assert resolved.patient_id == "95"
    assert request.patient_id == ""


def test_explicit_legacy_patient_id_is_preserved():
    request = NewPatientRequest(patient_id="IMPORT-001", display_name="ΝΕΟΣ")
    resolved = resolve_patient_id_request(
        request,
        existing_patients=[Patient("94", "A")],
    )
    assert resolved is request
    assert resolved.patient_id == "IMPORT-001"
