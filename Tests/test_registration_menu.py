from datetime import date

import pytest

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


def test_patient_form_declares_config_backed_room_and_status_fields():
    spec = registration_form_spec("new_patient")
    fields = {field.key: field for field in spec.fields}

    assert fields["patient_id"].required is True
    assert fields["display_name"].required is True
    assert fields["room"].source == "rooms"
    assert fields["status"].source == "patient_statuses"
    assert fields["infectious"].field_type is FormFieldType.BOOLEAN
    assert fields["infectious"].default is False


def test_therapist_form_does_not_offer_unsupported_capability_field():
    spec = registration_form_spec(RegistrationMenuAction.NEW_THERAPIST)

    assert [field.key for field in spec.fields] == ["display_name"]


def test_student_form_declares_supervisor_source_and_capability_defaults():
    spec = registration_form_spec("new_student")
    fields = {field.key: field for field in spec.fields}

    assert fields["supervisor_therapist_id"].source == "therapists"
    assert fields["replacement_capable"].default is True
    assert fields["robotic_capable"].default is False


def test_unknown_menu_action_stops_cleanly():
    with pytest.raises(ValueError, match="Unknown registration menu action"):
        registration_form_spec("delete_everything")


def test_build_patient_request_trims_text_and_parses_boolean():
    request = build_registration_request(
        "new_patient",
        {
            "patient_id": "  P-100  ",
            "display_name": "  Test Patient  ",
            "room": "  12A ",
            "infectious": "ναι",
            "status": " Active ",
        },
    )

    assert request == NewPatientRequest(
        patient_id="P-100",
        display_name="Test Patient",
        room="12A",
        infectious=True,
        status="Active",
    )


def test_build_patient_request_converts_blank_optional_text_to_none():
    request = build_registration_request(
        RegistrationMenuAction.NEW_PATIENT,
        {
            "patient_id": "P-101",
            "display_name": "Patient Two",
            "room": " ",
            "status": "",
        },
    )

    assert request.room is None
    assert request.status is None
    assert request.infectious is False


def test_build_therapist_request_keeps_unsupported_robotic_capability_false():
    request = build_registration_request(
        "new_therapist",
        {"display_name": " Therapist A "},
    )

    assert request == NewTherapistRequest(
        display_name="Therapist A",
        robotic_capable=False,
    )


def test_build_student_request_accepts_ui_date_formats_and_boolean_values():
    request = build_registration_request(
        "new_student",
        {
            "student_id": " S-7 ",
            "display_name": " Student Seven ",
            "student_number": "7",
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
        student_number=7,
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
            "student_number": 8,
            "placement_start": date(2026, 10, 1),
            "placement_end": date(2026, 11, 1),
        },
    )

    assert request.replacement_capable is True
    assert request.robotic_capable is False


@pytest.mark.parametrize(
    ("values", "message"),
    [
        ({"patient_id": "", "display_name": "Patient"}, "patient_id is required"),
        ({"patient_id": "P-1", "display_name": ""}, "display_name is required"),
    ],
)
def test_patient_required_fields_stop_before_backend(values, message):
    with pytest.raises(ValueError, match=message):
        build_registration_request("new_patient", values)


def test_student_invalid_integer_stops_cleanly():
    with pytest.raises(ValueError, match="student_number must be an integer"):
        build_registration_request(
            "new_student",
            {
                "student_id": "S-9",
                "display_name": "Student Nine",
                "student_number": "nine",
                "placement_start": "01/10/2026",
                "placement_end": "01/11/2026",
            },
        )


def test_student_invalid_date_stops_cleanly():
    with pytest.raises(ValueError, match="placement_start must use"):
        build_registration_request(
            "new_student",
            {
                "student_id": "S-9",
                "display_name": "Student Nine",
                "student_number": 9,
                "placement_start": "10.01.2026",
                "placement_end": "01/11/2026",
            },
        )
