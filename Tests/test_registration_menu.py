from datetime import date

import pytest

from rehab_core.models import PatientType
from rehab_core.registration import (
    NewPatientRequest,
    NewStudentRequest,
    NewTherapistRequest,
)
from rehab_excel.registration_menu import (
    FormFieldType,
    RegistrationMenuAction,
    build_registration_request,
    registration_form_spec,
    registration_menu_actions,
)
from rehab_excel.registration_orchestrator import RegistrationKind


def test_menu_exposes_patient_therapist_student_in_stable_order():
    specs = registration_menu_actions()

    assert [spec.action for spec in specs] == [
        RegistrationMenuAction.NEW_PATIENT,
        RegistrationMenuAction.NEW_THERAPIST,
        RegistrationMenuAction.NEW_STUDENT,
    ]
    assert [spec.kind for spec in specs] == [
        RegistrationKind.PATIENT,
        RegistrationKind.THERAPIST,
        RegistrationKind.STUDENT,
    ]


def test_patient_form_declares_auto_id_type_mrn_and_inpatient_fields():
    spec = registration_form_spec("new_patient")
    fields = {field.key: field for field in spec.fields}

    assert fields["patient_id"].required is False
    assert fields["patient_type"].required is True
    assert fields["patient_type"].default == PatientType.INPATIENT.value
    assert fields["patient_type"].source == "patient_types"
    assert fields["hospital_mrn"].required is False
    assert fields["display_name"].required is True
    assert fields["room"].source == "rooms"
    assert fields["status"].source == "patient_statuses"
    assert fields["infectious"].field_type is FormFieldType.BOOLEAN
    assert fields["infectious"].default is False


def test_therapist_form_does_not_offer_unsupported_capability_field():
    spec = registration_form_spec(RegistrationMenuAction.NEW_THERAPIST)
    assert [field.key for field in spec.fields] == ["display_name"]


def test_student_form_declares_single_id_supervisor_source_and_capability_defaults():
    spec = registration_form_spec("new_student")
    fields = {field.key: field for field in spec.fields}
    assert "student_number" not in fields
    assert fields["student_id"].required is True
    assert fields["supervisor_therapist_id"].source == "therapists"
    assert fields["replacement_capable"].default is True
    assert fields["robotic_capable"].default is False


def test_unknown_menu_action_stops_cleanly():
    with pytest.raises(ValueError, match="Unknown registration menu action"):
        registration_form_spec("delete_everything")


def test_build_patient_request_supports_auto_id_default_inpatient():
    request = build_registration_request(
        "new_patient",
        {
            "display_name": "  Test Patient  ",
            "room": "  12A ",
            "infectious": "ναι",
            "status": " Active ",
        },
    )

    assert request == NewPatientRequest(
        patient_id="",
        display_name="Test Patient",
        room="12A",
        infectious=True,
        status="Active",
        patient_type=PatientType.INPATIENT,
        hospital_mrn=None,
    )


def test_build_outpatient_request_accepts_greek_type_and_hospital_mrn():
    request = build_registration_request(
        RegistrationMenuAction.NEW_PATIENT,
        {
            "patient_id": "",
            "patient_type": "Εξωτερικός",
            "hospital_mrn": " MRN-77 ",
            "display_name": "Patient Two",
            "room": " ",
            "status": "",
        },
    )

    assert request.patient_id == ""
    assert request.patient_type is PatientType.OUTPATIENT
    assert request.hospital_mrn == "MRN-77"
    assert request.room is None
    assert request.status is None
    assert request.infectious is False


def test_build_patient_request_accepts_explicit_legacy_id():
    request = build_registration_request(
        "new_patient",
        {"patient_id": " P-101 ", "display_name": "Patient"},
    )
    assert request.patient_id == "P-101"


def test_invalid_patient_type_stops_cleanly():
    with pytest.raises(ValueError, match="patient_type must be"):
        build_registration_request(
            "new_patient",
            {"display_name": "Patient", "patient_type": "visitor"},
        )


def test_build_therapist_request_keeps_unsupported_robotic_capability_false():
    request = build_registration_request(
        "new_therapist",
        {"display_name": " Therapist A "},
    )
    assert request == NewTherapistRequest(display_name="Therapist A", robotic_capable=False)


def test_build_student_request_accepts_ui_date_formats_and_boolean_values():
    request = build_registration_request(
        "new_student",
        {
            "student_id": " S-7 ",
            "display_name": " Student Seven ",
            "placement_start": "01/10/2026",
            "placement_end": "2026-12-31",
            "supervisor_therapist_id": " Therapist A ",
            "replacement_capable": "όχι",
            "robotic_capable": 1,
        },
    )

    assert request == NewStudentRequest(
        student_id="S-7",
        display_name="Student Seven",
        placement_start=date(2026, 10, 1),
        placement_end=date(2026, 12, 31),
        supervisor_therapist_id="Therapist A",
        replacement_capable=False,
        robotic_capable=True,
    )


def test_build_student_request_uses_capability_defaults_when_omitted():
    request = build_registration_request(
        "new_student",
        {
            "student_id": "S-8",
            "display_name": "Student Eight",
            "placement_start": date(2026, 10, 1),
            "placement_end": date(2026, 11, 1),
        },
    )
    assert request.replacement_capable is True
    assert request.robotic_capable is False


def test_patient_name_is_the_only_required_text_field_at_ui_boundary():
    with pytest.raises(ValueError, match="display_name is required"):
        build_registration_request("new_patient", {"patient_id": ""})


def test_student_invalid_date_stops_cleanly():
    with pytest.raises(ValueError, match="placement_start must use"):
        build_registration_request(
            "new_student",
            {
                "student_id": "S-9",
                "display_name": "Student Nine",
                "placement_start": "10.01.2026",
                "placement_end": "01/11/2026",
            },
        )
