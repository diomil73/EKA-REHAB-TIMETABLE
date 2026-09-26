from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Protocol

from rehab_core.models import PatientType
from rehab_core.registration import NewPatientRequest, validate_new_patient

from .patient_registry_source import read_patient_registry
from .reader import read_settings


class PatientRegistrationWriteError(RuntimeError):
    """Raised when a safe patient-registration preview cannot be produced."""


PATIENT_TYPE_HEADERS = ("PatientType", "ΤύποςΑσθενή", "Τύπος Ασθενή")
HOSPITAL_MRN_HEADERS = ("HospitalMRN", "ΑΜ Νοσοκομείου", "ΑΜΝοσοκομείου")
OUTPATIENT_SCHEDULE_HEADERS = (
    "PatientID",
    "Ασθενής",
    "Θεραπεία",
    "Ώρα",
    "Ημέρες",
    "Θεραπευτής",
)


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _norm(value: object) -> str:
    return str(value or "").strip().casefold()


def _is_yes_token(value: object) -> bool:
    return _norm(value) in {"ν", "ναι", "n", "yes", "y", "true", "1"}


def resolve_infectious_cell_value(
    infectious: bool,
    configured_values: tuple[str, ...],
) -> str:
    """Use the workbook's own yes/no labels whenever possible."""

    if infectious:
        for value in configured_values:
            if _is_yes_token(value):
                return value
        return "Ν"

    for value in configured_values:
        if not _is_yes_token(value):
            return value
    return ""


def choose_patient_target_row(
    last_patient_id_row: int,
    last_patient_name_row: int,
    *,
    table_first_data_row: int | None = None,
    table_last_data_row: int | None = None,
) -> int:
    """Return the next logical PATIENTS row, not the physical end of a blank table."""

    logical_row = max(2, max(last_patient_id_row, last_patient_name_row, 1) + 1)
    if table_first_data_row is None or table_last_data_row is None:
        return logical_row
    if logical_row < table_first_data_row:
        return table_first_data_row
    if logical_row <= table_last_data_row:
        return logical_row
    return logical_row


def _remove_preview_file(path: Path, *, required: bool) -> None:
    if not path.exists():
        return
    try:
        path.unlink()
    except PermissionError as exc:
        if required:
            raise PatientRegistrationWriteError(
                "Preview workbook is currently open or locked by Excel. "
                f"Close it and retry: {path}"
            ) from exc


class PatientRegistrationBackend(Protocol):
    def append_patient(
        self,
        workbook_path: Path,
        request: NewPatientRequest,
        *,
        infectious_cell_value: str,
    ) -> int:
        """Append to PATIENTS and return the Excel row that was written."""
        ...


class Win32ComPatientRegistrationBackend:
    """Append one patient using Excel itself, on an already-created copy only."""

    @staticmethod
    def _last_used_row(ws, column: int) -> int:
        return int(ws.Cells(ws.Rows.Count, column).End(-4162).Row)

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
        physical_id_row = self._last_used_row(ws, 1)
        physical_name_row = self._last_used_row(ws, 3)
        scan_bottom = max(physical_id_row, physical_name_row, 1)

        last_id_row = self._last_nonblank_value_row(ws, 1, scan_bottom)
        last_name_row = self._last_nonblank_value_row(ws, 3, scan_bottom)
        logical_row = choose_patient_target_row(last_id_row, last_name_row)

        try:
            count = int(ws.ListObjects.Count)
        except Exception:
            count = 0
        for index in range(1, count + 1):
            table = ws.ListObjects(index)
            first_col = int(table.Range.Column)
            last_col = first_col + int(table.Range.Columns.Count) - 1
            if int(table.HeaderRowRange.Row) != 1 or first_col > 1 or last_col < 5:
                continue

            try:
                data_body = table.DataBodyRange
            except Exception:
                data_body = None

            if data_body is not None:
                first_data_row = int(data_body.Row)
                last_data_row = first_data_row + int(data_body.Rows.Count) - 1
                target = choose_patient_target_row(
                    last_id_row,
                    last_name_row,
                    table_first_data_row=first_data_row,
                    table_last_data_row=last_data_row,
                )
                if first_data_row <= target <= last_data_row:
                    return target
                if target == last_data_row + 1:
                    return int(table.ListRows.Add().Range.Row)

        return logical_row

    @staticmethod
    def _header_columns(ws) -> dict[str, int]:
        try:
            used_columns = int(ws.UsedRange.Columns.Count)
            first_column = int(ws.UsedRange.Column)
            last_column = first_column + used_columns - 1
        except Exception:
            last_column = 7
        last_column = max(last_column, 7)
        result: dict[str, int] = {}
        for column in range(1, last_column + 1):
            value = ws.Cells(1, column).Value2
            text = str(value or "").strip()
            if text:
                result[text] = column
        return result

    @staticmethod
    def _find_header(headers: dict[str, int], names: tuple[str, ...]) -> int | None:
        for name in names:
            if name in headers:
                return headers[name]
        return None

    def _ensure_patient_extension_columns(self, ws) -> tuple[int, int]:
        headers = self._header_columns(ws)
        type_col = self._find_header(headers, PATIENT_TYPE_HEADERS)
        mrn_col = self._find_header(headers, HOSPITAL_MRN_HEADERS)

        occupied = set(headers.values())
        next_col = max(occupied or {5}) + 1
        if type_col is None:
            while next_col in occupied:
                next_col += 1
            type_col = next_col
            ws.Cells(1, type_col).Value = "PatientType"
            occupied.add(type_col)
            next_col += 1
        if mrn_col is None:
            while next_col in occupied:
                next_col += 1
            mrn_col = next_col
            ws.Cells(1, mrn_col).Value = "HospitalMRN"

        return type_col, mrn_col

    @staticmethod
    def _ensure_outpatient_schedule_sheet(workbook) -> None:
        try:
            ws = workbook.Worksheets("OUTPATIENT_SCHEDULE")
        except Exception:
            ws = workbook.Worksheets.Add(
                After=workbook.Worksheets(workbook.Worksheets.Count)
            )
            ws.Name = "OUTPATIENT_SCHEDULE"

        for column, header in enumerate(OUTPATIENT_SCHEDULE_HEADERS, start=1):
            current = str(ws.Cells(1, column).Value2 or "").strip()
            if current and current != header:
                raise PatientRegistrationWriteError(
                    "OUTPATIENT_SCHEDULE exists but its header schema is incompatible"
                )
            if not current:
                ws.Cells(1, column).Value = header

    def append_patient(
        self,
        workbook_path: Path,
        request: NewPatientRequest,
        *,
        infectious_cell_value: str,
    ) -> int:
        if sys.platform != "win32":
            raise PatientRegistrationWriteError(
                "Patient preview write-back requires Windows with Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise PatientRegistrationWriteError(
                "pywin32 is required for patient preview write-back"
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
                ws = workbook.Worksheets("PATIENTS")
            except Exception as exc:
                raise PatientRegistrationWriteError("Workbook has no PATIENTS sheet") from exc

            type_col, mrn_col = self._ensure_patient_extension_columns(ws)
            self._ensure_outpatient_schedule_sheet(workbook)
            target_row = self._target_row(ws)

            if target_row > 2:
                try:
                    target = ws.Range(f"A{target_row}:E{target_row}")
                    if target.ListObject is None:
                        ws.Range(f"A{target_row - 1}:E{target_row - 1}").Copy()
                        target.PasteSpecial(Paste=-4122)
                except Exception:
                    pass

            is_outpatient = request.patient_type == PatientType.OUTPATIENT
            ws.Cells(target_row, 1).Value = request.patient_id.strip()
            ws.Cells(target_row, 2).Value = "" if is_outpatient else (request.room or "").strip()
            ws.Cells(target_row, 3).Value = request.display_name.strip()
            ws.Cells(target_row, 4).Value = "" if is_outpatient else infectious_cell_value
            ws.Cells(target_row, 5).Value = "" if is_outpatient else (request.status or "").strip()
            ws.Cells(target_row, type_col).Value = (
                "Εξωτερικός" if is_outpatient else "Εσωτερικός"
            )
            ws.Cells(target_row, mrn_col).Value = (request.hospital_mrn or "").strip()
            workbook.Application.CutCopyMode = False
            workbook.Save()
            return target_row
        except PatientRegistrationWriteError:
            raise
        except Exception as exc:
            raise PatientRegistrationWriteError(
                f"Excel patient registration failed: {exc}"
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
class PatientRegistrationPreviewReport:
    source_path: str
    output_path: str
    patient_id: str
    excel_row: int
    source_unchanged: bool
    verified_in_output: bool


def create_patient_registration_preview(
    source_path: str | Path,
    output_path: str | Path,
    request: NewPatientRequest,
    *,
    backend: PatientRegistrationBackend | None = None,
    overwrite: bool = False,
) -> PatientRegistrationPreviewReport:
    """Register one patient in a NEW .xlsm copy and verify the result."""

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise PatientRegistrationWriteError(f"Source workbook not found: {source}")
    if source == output:
        raise PatientRegistrationWriteError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise PatientRegistrationWriteError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise PatientRegistrationWriteError(f"Output already exists: {output}")

    existing = read_patient_registry(source)
    settings = read_settings(source)
    check = validate_new_patient(
        request,
        existing_patients=existing,
        allowed_statuses=settings.patient_statuses,
        allowed_rooms=settings.rooms,
    )
    if not check.allowed:
        detail = "; ".join(f"{issue.field}: {issue.message}" for issue in check.issues)
        raise PatientRegistrationWriteError(f"Patient validation failed: {detail}")

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        _remove_preview_file(output, required=True)
    shutil.copy2(source, output)

    selected_backend = backend or Win32ComPatientRegistrationBackend()
    infectious_value = resolve_infectious_cell_value(
        request.infectious,
        settings.yes_no_values,
    )
    try:
        excel_row = selected_backend.append_patient(
            output,
            request,
            infectious_cell_value=infectious_value,
        )
    except Exception:
        _remove_preview_file(output, required=False)
        raise

    source_after = _sha256(source)
    if source_before != source_after:
        _remove_preview_file(output, required=False)
        raise PatientRegistrationWriteError("Source workbook changed during preview creation")

    matches = [
        patient
        for patient in read_patient_registry(output)
        if str(patient.patient_id).strip().casefold()
        == request.patient_id.strip().casefold()
    ]
    expected_mrn = (request.hospital_mrn or "").strip() or None
    verified = (
        len(matches) == 1
        and matches[0].display_name.strip() == request.display_name.strip()
        and matches[0].patient_type == request.patient_type
        and matches[0].hospital_mrn == expected_mrn
    )
    if not verified:
        _remove_preview_file(output, required=False)
        raise PatientRegistrationWriteError(
            "Preview verification failed: patient metadata did not read back exactly"
        )

    return PatientRegistrationPreviewReport(
        source_path=str(source),
        output_path=str(output),
        patient_id=request.patient_id.strip(),
        excel_row=excel_row,
        source_unchanged=True,
        verified_in_output=True,
    )
