from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Callable

from rehab_core.registration import (
    NewPatientRequest,
    NewStudentRequest,
    NewTherapistRequest,
)

from .patient_registration import (
    PatientRegistrationPreviewReport,
    PatientRegistrationWriteError,
    create_patient_registration_preview,
)
from .student_registration import (
    StudentRegistrationPreviewReport,
    StudentRegistrationWriteError,
    create_student_registration_preview,
)
from .therapist_registration import (
    TherapistRegistrationPreviewReport,
    TherapistRegistrationWriteError,
    create_therapist_registration_preview,
)


class RegistrationKind(str, Enum):
    PATIENT = "patient"
    THERAPIST = "therapist"
    STUDENT = "student"


RegistrationRequest = NewPatientRequest | NewTherapistRequest | NewStudentRequest
RegistrationReport = (
    PatientRegistrationPreviewReport
    | TherapistRegistrationPreviewReport
    | StudentRegistrationPreviewReport
)


PREVIEW_FILENAMES: dict[RegistrationKind, str] = {
    RegistrationKind.PATIENT: "NEW_PATIENT_PREVIEW.xlsm",
    RegistrationKind.THERAPIST: "NEW_THERAPIST_PREVIEW.xlsm",
    RegistrationKind.STUDENT: "NEW_STUDENT_PREVIEW.xlsm",
}


class RegistrationOrchestrationError(RuntimeError):
    """Uniform safety error for the future menu/form registration layer."""


@dataclass(frozen=True)
class RegistrationPreviewResult:
    kind: RegistrationKind
    subject_key: str
    display_name: str
    output_path: str
    excel_row: int
    source_unchanged: bool
    verified_in_output: bool
    schema_created: bool | None = None


@dataclass(frozen=True)
class RegistrationPreviewServices:
    """Injectable preview services keep orchestration testable without Excel COM."""

    patient: Callable[..., PatientRegistrationPreviewReport] = create_patient_registration_preview
    therapist: Callable[..., TherapistRegistrationPreviewReport] = create_therapist_registration_preview
    student: Callable[..., StudentRegistrationPreviewReport] = create_student_registration_preview


def registration_kind(request: RegistrationRequest) -> RegistrationKind:
    if isinstance(request, NewPatientRequest):
        return RegistrationKind.PATIENT
    if isinstance(request, NewTherapistRequest):
        return RegistrationKind.THERAPIST
    if isinstance(request, NewStudentRequest):
        return RegistrationKind.STUDENT
    raise TypeError(f"Unsupported registration request type: {type(request).__name__}")


def registration_preview_path(
    preview_dir: str | Path,
    request: RegistrationRequest,
) -> Path:
    kind = registration_kind(request)
    return Path(preview_dir) / PREVIEW_FILENAMES[kind]


def _normalize_result(
    request: RegistrationRequest,
    report: RegistrationReport,
) -> RegistrationPreviewResult:
    kind = registration_kind(request)

    if kind is RegistrationKind.PATIENT:
        assert isinstance(request, NewPatientRequest)
        assert isinstance(report, PatientRegistrationPreviewReport)
        return RegistrationPreviewResult(
            kind=kind,
            subject_key=request.patient_id.strip(),
            display_name=request.display_name.strip(),
            output_path=report.output_path,
            excel_row=report.excel_row,
            source_unchanged=report.source_unchanged,
            verified_in_output=report.verified_in_output,
        )

    if kind is RegistrationKind.THERAPIST:
        assert isinstance(request, NewTherapistRequest)
        assert isinstance(report, TherapistRegistrationPreviewReport)
        return RegistrationPreviewResult(
            kind=kind,
            subject_key=request.display_name.strip(),
            display_name=request.display_name.strip(),
            output_path=report.output_path,
            excel_row=report.excel_row,
            source_unchanged=report.source_unchanged,
            verified_in_output=report.verified_in_output,
        )

    assert isinstance(request, NewStudentRequest)
    assert isinstance(report, StudentRegistrationPreviewReport)
    return RegistrationPreviewResult(
        kind=kind,
        subject_key=request.student_id.strip(),
        display_name=request.display_name.strip(),
        output_path=report.output_path,
        excel_row=report.excel_row,
        source_unchanged=report.source_unchanged,
        verified_in_output=report.verified_in_output,
        schema_created=report.schema_created,
    )


def create_registration_preview(
    source_path: str | Path,
    preview_dir: str | Path,
    request: RegistrationRequest,
    *,
    overwrite: bool = False,
    services: RegistrationPreviewServices | None = None,
) -> RegistrationPreviewResult:
    """Route one formal registration through its proven safe preview workflow.

    This is the single backend entry point intended for the future Excel menu or
    Windows UI. It deliberately creates preview copies only. The three existing
    entity-specific workflows remain responsible for validation, COM writeback,
    source hashing, and read-back verification.
    """

    kind = registration_kind(request)
    output_path = registration_preview_path(preview_dir, request)
    selected = services or RegistrationPreviewServices()

    try:
        if kind is RegistrationKind.PATIENT:
            report = selected.patient(
                source_path,
                output_path,
                request,
                overwrite=overwrite,
            )
        elif kind is RegistrationKind.THERAPIST:
            report = selected.therapist(
                source_path,
                output_path,
                request,
                overwrite=overwrite,
            )
        else:
            report = selected.student(
                source_path,
                output_path,
                request,
                overwrite=overwrite,
            )
    except (
        PatientRegistrationWriteError,
        TherapistRegistrationWriteError,
        StudentRegistrationWriteError,
    ) as exc:
        raise RegistrationOrchestrationError(
            f"{kind.value} registration preview failed: {exc}"
        ) from exc

    return _normalize_result(request, report)
