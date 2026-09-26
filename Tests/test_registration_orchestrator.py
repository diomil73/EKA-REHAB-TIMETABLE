from datetime import date

import pytest

from rehab_core.registration import (
    NewPatientRequest,
    NewStudentRequest,
    NewTherapistRequest,
)
from rehab_excel.patient_registration import (
    PatientRegistrationPreviewReport,
    PatientRegistrationWriteError,
)
from rehab_excel.registration_orchestrator import (
    RegistrationKind,
    RegistrationOrchestrationError,
    RegistrationPreviewServices,
    create_registration_preview,
    registration_kind,
    registration_preview_path,
)
from rehab_excel.student_registration import StudentRegistrationPreviewReport
from rehab_excel.therapist_registration import TherapistRegistrationPreviewReport


def _patient_request():
    return NewPatientRequest("P-NEW", "ΝΕΟΣ ΑΣΘΕΝΗΣ")


def _therapist_request():
    return NewTherapistRequest("ΝΕΟΣ ΘΕΡΑΠΕΥΤΗΣ")


def _student_request():
    return NewStudentRequest(
        student_id="STU-NEW",
        display_name="ΝΕΟΣ ΦΟΙΤΗΤΗΣ",
        student_number=7,
        placement_start=date(2026, 10, 1),
        placement_end=date(2026, 12, 31),
    )


def test_registration_kind_and_preview_filename_are_entity_specific(tmp_path):
    patient = _patient_request()
    therapist = _therapist_request()
    student = _student_request()

    assert registration_kind(patient) is RegistrationKind.PATIENT
    assert registration_kind(therapist) is RegistrationKind.THERAPIST
    assert registration_kind(student) is RegistrationKind.STUDENT
    assert registration_preview_path(tmp_path, patient).name == "NEW_PATIENT_PREVIEW.xlsm"
    assert registration_preview_path(tmp_path, therapist).name == "NEW_THERAPIST_PREVIEW.xlsm"
    assert registration_preview_path(tmp_path, student).name == "NEW_STUDENT_PREVIEW.xlsm"


def test_patient_registration_is_routed_and_normalized(tmp_path):
    calls = []

    def patient_service(source, output, request, *, overwrite):
        calls.append((source, output, request, overwrite))
        return PatientRegistrationPreviewReport(
            source_path=str(source),
            output_path=str(output),
            patient_id=request.patient_id,
            excel_row=99,
            source_unchanged=True,
            verified_in_output=True,
        )

    services = RegistrationPreviewServices(
        patient=patient_service,
        therapist=lambda *args, **kwargs: pytest.fail("wrong therapist route"),
        student=lambda *args, **kwargs: pytest.fail("wrong student route"),
    )
    request = _patient_request()

    result = create_registration_preview(
        "baseline.xlsm",
        tmp_path,
        request,
        overwrite=True,
        services=services,
    )

    assert result.kind is RegistrationKind.PATIENT
    assert result.subject_key == "P-NEW"
    assert result.display_name == "ΝΕΟΣ ΑΣΘΕΝΗΣ"
    assert result.excel_row == 99
    assert result.schema_created is None
    assert calls[0][1].name == "NEW_PATIENT_PREVIEW.xlsm"
    assert calls[0][3] is True


def test_therapist_registration_is_routed_and_normalized(tmp_path):
    def therapist_service(source, output, request, *, overwrite):
        return TherapistRegistrationPreviewReport(
            source_path=str(source),
            output_path=str(output),
            therapist_name=request.display_name,
            excel_row=19,
            source_unchanged=True,
            verified_in_output=True,
            other_settings_unchanged=True,
        )

    services = RegistrationPreviewServices(
        patient=lambda *args, **kwargs: pytest.fail("wrong patient route"),
        therapist=therapist_service,
        student=lambda *args, **kwargs: pytest.fail("wrong student route"),
    )

    result = create_registration_preview(
        "baseline.xlsm",
        tmp_path,
        _therapist_request(),
        services=services,
    )

    assert result.kind is RegistrationKind.THERAPIST
    assert result.subject_key == "ΝΕΟΣ ΘΕΡΑΠΕΥΤΗΣ"
    assert result.excel_row == 19
    assert result.schema_created is None


def test_student_registration_preserves_schema_created_flag(tmp_path):
    def student_service(source, output, request, *, overwrite):
        return StudentRegistrationPreviewReport(
            source_path=str(source),
            output_path=str(output),
            student_id=request.student_id,
            excel_row=2,
            schema_created=True,
            source_unchanged=True,
            verified_in_output=True,
        )

    services = RegistrationPreviewServices(
        patient=lambda *args, **kwargs: pytest.fail("wrong patient route"),
        therapist=lambda *args, **kwargs: pytest.fail("wrong therapist route"),
        student=student_service,
    )

    result = create_registration_preview(
        "baseline.xlsm",
        tmp_path,
        _student_request(),
        services=services,
    )

    assert result.kind is RegistrationKind.STUDENT
    assert result.subject_key == "STU-NEW"
    assert result.excel_row == 2
    assert result.schema_created is True


def test_entity_write_errors_become_one_menu_facing_safety_error(tmp_path):
    def failing_patient(*args, **kwargs):
        raise PatientRegistrationWriteError("locked preview")

    services = RegistrationPreviewServices(
        patient=failing_patient,
        therapist=lambda *args, **kwargs: pytest.fail("wrong therapist route"),
        student=lambda *args, **kwargs: pytest.fail("wrong student route"),
    )

    with pytest.raises(
        RegistrationOrchestrationError,
        match="patient registration preview failed: locked preview",
    ):
        create_registration_preview(
            "baseline.xlsm",
            tmp_path,
            _patient_request(),
            services=services,
        )


def test_unsupported_registration_request_is_rejected():
    with pytest.raises(TypeError, match="Unsupported registration request type"):
        registration_kind(object())  # type: ignore[arg-type]
