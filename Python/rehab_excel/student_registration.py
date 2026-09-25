from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Protocol
import unicodedata

from rehab_core.models import Student
from rehab_core.registration import NewStudentRequest, validate_new_student

from .reader import read_settings
from .student_registry import (
    STUDENT_REGISTRY_HEADERS,
    STUDENT_REGISTRY_SHEET,
    StudentRegistryError,
    read_students,
    validate_student_registry_headers,
)
from .student_registry_schema import (
    StudentRegistrySchemaBackend,
    Win32ComStudentRegistrySchemaBackend,
)


class StudentRegistrationWriteError(RuntimeError):
    """Raised when a safe student-registration preview cannot be produced."""


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _key(value: object) -> str:
    text = unicodedata.normalize("NFD", str(value).strip().casefold())
    return "".join(char for char in text if not unicodedata.combining(char))


def _is_yes_token(value: object) -> bool:
    return _key(value) in {"ν", "ναι", "yes", "y", "true", "1"}


def _is_no_token(value: object) -> bool:
    return _key(value) in {"οχι", "no", "n", "false", "0"}


def resolve_boolean_cell_value(
    value: bool,
    configured_values: tuple[str, ...],
) -> str:
    """Use the workbook's configured yes/no labels when possible."""

    matcher = _is_yes_token if value else _is_no_token
    for item in configured_values:
        if matcher(item):
            return item
    return "ΝΑΙ" if value else "ΟΧΙ"


def choose_student_target_row(
    last_student_id_row: int,
    last_student_name_row: int,
    last_student_number_row: int,
) -> int:
    """Return the next logical STUDENTS data row."""

    return max(
        2,
        max(
            int(last_student_id_row),
            int(last_student_name_row),
            int(last_student_number_row),
            1,
        )
        + 1,
    )


def _student_mismatch_detail(
    expected: list[Student],
    actual: list[Student],
    *,
    target_student_id: str,
) -> str:
    """Return a compact read-back diagnostic without dumping unrelated rows."""

    target_key = _key(target_student_id)
    expected_target = next(
        (student for student in expected if _key(student.student_id) == target_key),
        None,
    )
    actual_target = next(
        (student for student in actual if _key(student.student_id) == target_key),
        None,
    )

    if actual_target is None:
        actual_ids = [student.student_id for student in actual]
        return (
            f"target StudentID {target_student_id!r} was not found after read-back; "
            f"actual StudentIDs={actual_ids!r}"
        )
    if expected_target is None:
        return "internal verification error: expected target student is missing"

    fields = (
        "student_id",
        "display_name",
        "student_number",
        "placement_start",
        "placement_end",
        "supervisor_therapist_id",
        "replacement_capable",
        "robotic_capable",
        "max_daily_timeslots",
    )
    differences = []
    for field in fields:
        expected_value = getattr(expected_target, field)
        actual_value = getattr(actual_target, field)
        if expected_value != actual_value:
            differences.append(
                f"{field}: expected={expected_value!r}, actual={actual_value!r}"
            )

    if differences:
        return "; ".join(differences)
    return (
        f"target student matches, but registry sequence/count differs: "
        f"expected_count={len(expected)}, actual_count={len(actual)}"
    )


def _remove_preview(path: Path, *, required: bool) -> None:
    if not path.exists():
        return
    try:
        path.unlink()
    except PermissionError as exc:
        if required:
            raise StudentRegistrationWriteError(
                "Preview workbook is currently open or locked by Excel. "
                f"Close it and retry: {path}"
            ) from exc


class StudentRegistrationBackend(Protocol):
    def append_student(
        self,
        workbook_path: Path,
        request: NewStudentRequest,
        *,
        replacement_cell_value: str,
        robotic_cell_value: str,
    ) -> int:
        """Append one student and return the Excel row written."""
        ...


class Win32ComStudentRegistrationBackend:
    """Append one validated student to STUDENTS on an already-created copy."""

    @staticmethod
    def _last_used_row(ws, column: int) -> int:
        return int(ws.Cells(ws.Rows.Count, column).End(-4162).Row)  # xlUp

    @staticmethod
    def _last_nonblank_value_row(ws, column: int, upper_row: int) -> int:
        upper_row = max(1, int(upper_row))
        values = ws.Range(ws.Cells(1, column), ws.Cells(upper_row, column)).Value2
        if upper_row == 1:
            sequence = ((values,),)
        else:
            sequence = values

        last = 1
        for row_index, item in enumerate(sequence, start=1):
            value = item[0] if isinstance(item, tuple) else item
            if value is None:
                continue
            if isinstance(value, str) and not value.strip():
                continue
            last = row_index
        return last

    def _target_row(self, ws) -> int:
        upper = max(
            self._last_used_row(ws, 1),
            self._last_used_row(ws, 2),
            self._last_used_row(ws, 3),
            1,
        )
        return choose_student_target_row(
            self._last_nonblank_value_row(ws, 1, upper),
            self._last_nonblank_value_row(ws, 2, upper),
            self._last_nonblank_value_row(ws, 3, upper),
        )

    def append_student(
        self,
        workbook_path: Path,
        request: NewStudentRequest,
        *,
        replacement_cell_value: str,
        robotic_cell_value: str,
    ) -> int:
        if sys.platform != "win32":
            raise StudentRegistrationWriteError(
                "Student preview write-back requires Windows with Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise StudentRegistrationWriteError(
                "pywin32 is required for student preview write-back"
            ) from exc

        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            workbook = excel.Workbooks.Open(
                str(workbook_path.resolve()),
                UpdateLinks=0,
                ReadOnly=False,
            )
            try:
                ws = workbook.Worksheets(STUDENT_REGISTRY_SHEET)
            except Exception as exc:
                raise StudentRegistrationWriteError(
                    "Workbook has no STUDENTS sheet after schema preparation"
                ) from exc

            actual_headers = tuple(
                str(ws.Cells(1, index).Value2 or "").strip()
                for index in range(1, len(STUDENT_REGISTRY_HEADERS) + 1)
            )
            try:
                validate_student_registry_headers(actual_headers)
            except StudentRegistryError as exc:
                raise StudentRegistrationWriteError(str(exc)) from exc

            target_row = self._target_row(ws)
            ws.Cells(target_row, 1).Value = request.student_id.strip()
            ws.Cells(target_row, 2).Value = request.display_name.strip()
            ws.Cells(target_row, 3).Value = int(request.student_number)
            ws.Cells(target_row, 4).Value = datetime.combine(request.placement_start, time.min)
            ws.Cells(target_row, 5).Value = datetime.combine(request.placement_end, time.min)
            ws.Cells(target_row, 6).Value = (request.supervisor_therapist_id or "").strip()
            ws.Cells(target_row, 7).Value = replacement_cell_value
            ws.Cells(target_row, 8).Value = robotic_cell_value

            try:
                ws.Range(f"D{target_row}:E{target_row}").NumberFormat = "dd/mm/yyyy"
            except Exception:
                pass

            ws.Columns("A:H").AutoFit()
            workbook.Save()
            return target_row
        except StudentRegistrationWriteError:
            raise
        except Exception as exc:
            raise StudentRegistrationWriteError(
                f"Excel student registration failed: {exc}"
            ) from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass


@dataclass(frozen=True)
class StudentRegistrationPreviewReport:
    source_path: str
    output_path: str
    student_id: str
    excel_row: int
    schema_created: bool
    source_unchanged: bool
    verified_in_output: bool


def create_student_registration_preview(
    source_path: str | Path,
    output_path: str | Path,
    request: NewStudentRequest,
    *,
    schema_backend: StudentRegistrySchemaBackend | None = None,
    backend: StudentRegistrationBackend | None = None,
    overwrite: bool = False,
) -> StudentRegistrationPreviewReport:
    """Create a copied .xlsm, ensure STUDENTS exists, append and verify one student."""

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise StudentRegistrationWriteError(f"Source workbook not found: {source}")
    if source == output:
        raise StudentRegistrationWriteError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise StudentRegistrationWriteError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise StudentRegistrationWriteError(f"Output already exists: {output}")

    try:
        existing_students = read_students(source)
    except StudentRegistryError as exc:
        raise StudentRegistrationWriteError(
            f"Existing STUDENTS registry is invalid: {exc}"
        ) from exc
    settings = read_settings(source)
    check = validate_new_student(
        request,
        existing_students=existing_students,
        known_therapist_ids=settings.therapist_names,
    )
    if not check.allowed:
        detail = "; ".join(f"{issue.field}: {issue.message}" for issue in check.issues)
        raise StudentRegistrationWriteError(f"Student validation failed: {detail}")

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        _remove_preview(output, required=True)
    shutil.copy2(source, output)

    selected_schema_backend = schema_backend or Win32ComStudentRegistrySchemaBackend()
    selected_backend = backend or Win32ComStudentRegistrationBackend()
    replacement_value = resolve_boolean_cell_value(
        request.replacement_capable,
        settings.yes_no_values,
    )
    robotic_value = resolve_boolean_cell_value(
        request.robotic_capable,
        settings.yes_no_values,
    )

    try:
        schema_created = selected_schema_backend.ensure_schema(output)
        excel_row = selected_backend.append_student(
            output,
            request,
            replacement_cell_value=replacement_value,
            robotic_cell_value=robotic_value,
        )
    except Exception:
        _remove_preview(output, required=False)
        raise

    source_after = _sha256(source)
    if source_before != source_after:
        _remove_preview(output, required=False)
        raise StudentRegistrationWriteError(
            "Source workbook changed during student registration preview"
        )

    try:
        output_students = read_students(output)
    except StudentRegistryError as exc:
        _remove_preview(output, required=False)
        raise StudentRegistrationWriteError(
            f"Student preview verification failed: {exc}"
        ) from exc

    expected_student = Student(
        student_id=request.student_id.strip(),
        display_name=request.display_name.strip(),
        student_number=request.student_number,
        placement_start=request.placement_start,
        placement_end=request.placement_end,
        supervisor_therapist_id=(request.supervisor_therapist_id or "").strip() or None,
        replacement_capable=request.replacement_capable,
        robotic_capable=request.robotic_capable,
    )
    expected_students = [*existing_students, expected_student]
    normalized_matches = sum(
        1 for student in output_students if _key(student.student_id) == _key(request.student_id)
    )
    verified = output_students == expected_students and normalized_matches == 1
    if not verified:
        detail = _student_mismatch_detail(
            expected_students,
            output_students,
            target_student_id=request.student_id,
        )
        _remove_preview(output, required=False)
        raise StudentRegistrationWriteError(
            "Preview verification failed: " + detail
        )

    return StudentRegistrationPreviewReport(
        source_path=str(source),
        output_path=str(output),
        student_id=request.student_id.strip(),
        excel_row=excel_row,
        schema_created=schema_created,
        source_unchanged=True,
        verified_in_output=True,
    )
