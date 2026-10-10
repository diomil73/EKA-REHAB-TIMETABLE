from __future__ import annotations

from dataclasses import replace
from pathlib import Path
import shutil
import tempfile
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
    """Create a safe student preview with automatic sequential StudentID.

    When an overwrite preview already exists, treat that preview as the current
    working registry so repeated form submissions accumulate safely in the same
    preview instead of restarting the sequence from the unchanged source workbook.
    The authoritative source workbook is never modified.
    """

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    working_source = source
    temporary_source: Path | None = None

    if overwrite and output.exists():
        output.parent.mkdir(parents=True, exist_ok=True)
        handle = tempfile.NamedTemporaryFile(
            prefix="eka_student_preview_source_",
            suffix=".xlsm",
            dir=output.parent,
            delete=False,
        )
        handle.close()
        temporary_source = Path(handle.name)
        shutil.copy2(output, temporary_source)
        working_source = temporary_source

    try:
        existing = read_students(working_source)
        resolved = resolve_student_id_request(request, existing_students=existing)
        report = create_student_registration_preview(
            working_source,
            output,
            resolved,
            backend=backend,
            overwrite=overwrite,
        )
        return replace(
            report,
            source_path=str(source),
            source_unchanged=True,
        )
    finally:
        if temporary_source is not None and temporary_source.exists():
            try:
                temporary_source.unlink()
            except OSError:
                pass
