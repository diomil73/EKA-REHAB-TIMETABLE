from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Iterable

from rehab_core.models import Student
from rehab_core.registration import NewStudentRequest
from rehab_core.student_ids import next_sequential_student_id

from .student_registration import (
    StudentRegistrationBackend,
    StudentRegistrationPreviewReport,
    create_student_registration_preview,
)
from .student_registry import read_students


def resolve_student_id_request(
    request: NewStudentRequest,
    *,
    existing_students: Iterable[Student],
) -> NewStudentRequest:
    """Fill a blank StudentID with the next sequential internal ID.

    Explicit IDs remain supported for legacy/import tooling. Interactive UI
    callers should leave ``student_id`` blank so allocation is authoritative.
    """

    if request.student_id.strip():
        return request
    return replace(
        request,
        student_id=next_sequential_student_id(existing_students),
    )


def create_auto_student_registration_preview(
    source_path: str | Path,
    output_path: str | Path,
    request: NewStudentRequest,
    *,
    backend: StudentRegistrationBackend | None = None,
    overwrite: bool = False,
) -> StudentRegistrationPreviewReport:
    existing = read_students(source_path)
    resolved = resolve_student_id_request(request, existing_students=existing)
    return create_student_registration_preview(
        source_path,
        output_path,
        resolved,
        backend=backend,
        overwrite=overwrite,
    )
