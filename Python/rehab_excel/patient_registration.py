from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Protocol

from rehab_core.registration import NewPatientRequest, validate_new_patient

from .reader import read_patients, read_settings


class PatientRegistrationWriteError(RuntimeError):
    """Raised when a safe patient-registration preview cannot be produced."""


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
    """Use the workbook's own yes/no labels whenever possible.

    If the current workbook does not expose a usable yes/no list, True falls
    back to ``Ν`` (already understood by the reader) and False remains blank.
    """

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
    """Remove an old/failed preview without leaking raw WinError 32 tracebacks."""

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
        # Best-effort cleanup after another failure. Do not hide the original
        # exception merely because Excel still owns a handle to the preview.


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
        # xlUp = -4162. This reports formulas as used even when they display
        # blank, so it is only an upper bound for the value scan below.
        return int(ws.Cells(ws.Rows.Count, column).End(-4162).Row)

    @staticmethod
    def _last_nonblank_value_row(ws, column: int, upper_row: int) -> int:
        """Find the last row whose calculated/display value is genuinely nonblank.

        PATIENTS contains a pre-sized table. Cells below the visible patient
        list may contain formulas or table structure that make End(xlUp) report
        row 550 even though the last real patient is around row 98. Reading
        Value2 in one block lets us ignore formulas that currently evaluate to
        an empty string.
        """

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

        # PATIENTS may be a pre-sized Excel table with many blank/formula rows.
        # If the logical next patient row is already inside that table, write
        # directly there instead of extending the table at its physical bottom.
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

            target_row = self._target_row(ws)

            # For a plain range, borrow formats from the previous data row. If
            # the target is already inside a ListObject, the table owns styling.
            if target_row > 2:
                try:
                    target = ws.Range(f"A{target_row}:E{target_row}")
                    if target.ListObject is None:
                        ws.Range(f"A{target_row - 1}:E{target_row - 1}").Copy()
                        target.PasteSpecial(Paste=-4122)  # xlPasteFormats
                except Exception:
                    # Formatting is secondary to data safety. Never copy values
                    # from the previous row merely to obtain a style.
                    pass

            ws.Cells(target_row, 1).Value = request.patient_id.strip()
            ws.Cells(target_row, 2).Value = (request.room or "").strip()
            ws.Cells(target_row, 3).Value = request.display_name.strip()
            ws.Cells(target_row, 4).Value = infectious_cell_value
            ws.Cells(target_row, 5).Value = (request.status or "").strip()
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
    """Register one patient in a NEW .xlsm copy and verify the result.

    PATIENTS is authoritative, but the baseline workbook is never opened for
    writing. PATIENT_PLANNER is intentionally untouched in this first stage.
    """

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

    existing = read_patients(source)
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
        for patient in read_patients(output)
        if str(patient.patient_id).strip().casefold()
        == request.patient_id.strip().casefold()
    ]
    verified = len(matches) == 1 and matches[0].display_name.strip() == request.display_name.strip()
    if not verified:
        _remove_preview_file(output, required=False)
        raise PatientRegistrationWriteError(
            "Preview verification failed: new patient was not read back exactly once"
        )

    return PatientRegistrationPreviewReport(
        source_path=str(source),
        output_path=str(output),
        patient_id=request.patient_id.strip(),
        excel_row=excel_row,
        source_unchanged=True,
        verified_in_output=True,
    )
