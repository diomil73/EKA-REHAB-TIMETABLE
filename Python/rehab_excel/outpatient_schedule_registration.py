from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Iterable, Protocol

from rehab_core.day_patterns import parse_day_pattern
from rehab_core.models import BaseScheduleEntry, Patient

from .outpatient_presentation import validate_outpatient_base_entries
from .outpatient_schedule_source import (
    OUTPATIENT_REQUIRED_HEADERS,
    OUTPATIENT_SCHEDULE_SHEET,
    read_outpatient_schedule,
)
from .patient_registry_source import read_patient_registry


class OutpatientScheduleWriteError(RuntimeError):
    """Raised when a safe outpatient recurring-schedule preview cannot be produced."""


@dataclass(frozen=True)
class OutpatientScheduleRequest:
    patient_id: str
    treatment: str
    start_time: time
    day_pattern: str
    therapist_id: str | None = None
    target_base_entry_id: str | None = None


@dataclass(frozen=True)
class OutpatientSchedulePreviewReport:
    source_path: str
    output_path: str
    base_entry_id: str
    excel_row: int
    source_unchanged: bool
    verified_in_output: bool
    created: bool


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _target_row_from_entry_id(value: str) -> int:
    prefix = "outpatient:"
    if not value.startswith(prefix):
        raise OutpatientScheduleWriteError(
            f"Invalid outpatient base_entry_id: {value!r}"
        )
    try:
        row = int(value[len(prefix):])
    except ValueError as exc:
        raise OutpatientScheduleWriteError(
            f"Invalid outpatient base_entry_id: {value!r}"
        ) from exc
    if row < 2:
        raise OutpatientScheduleWriteError(
            f"Invalid outpatient schedule row in {value!r}"
        )
    return row


def validate_outpatient_schedule_request(
    request: OutpatientScheduleRequest,
    *,
    patients: Iterable[Patient],
) -> Patient:
    patient_id = request.patient_id.strip()
    treatment = request.treatment.strip()
    day_pattern = request.day_pattern.strip()
    if not patient_id:
        raise OutpatientScheduleWriteError("PatientID is required")
    if not treatment:
        raise OutpatientScheduleWriteError("Treatment is required")
    if not day_pattern:
        raise OutpatientScheduleWriteError("Day pattern is required")

    patient_by_id = {patient.patient_id: patient for patient in patients}
    patient = patient_by_id.get(patient_id)
    if patient is None:
        raise OutpatientScheduleWriteError(f"Unknown PatientID {patient_id!r}")
    if not patient.is_outpatient:
        raise OutpatientScheduleWriteError(
            f"Patient {patient_id!r} is not registered as outpatient"
        )

    try:
        parse_day_pattern(day_pattern)
    except ValueError as exc:
        raise OutpatientScheduleWriteError(
            f"Invalid outpatient day pattern {day_pattern!r}"
        ) from exc

    provisional = BaseScheduleEntry(
        base_entry_id=request.target_base_entry_id or "outpatient:new",
        patient_id=patient_id,
        treatment=treatment,
        start_time=request.start_time,
        day_pattern=day_pattern,
        therapist_id=(request.therapist_id or "").strip() or None,
        robotic=False,
    )
    try:
        validate_outpatient_base_entries((provisional,), patients)
    except Exception as exc:
        raise OutpatientScheduleWriteError(str(exc)) from exc
    return patient


class OutpatientScheduleBackend(Protocol):
    def upsert(
        self,
        workbook_path: Path,
        request: OutpatientScheduleRequest,
        *,
        patient_name: str,
    ) -> tuple[int, bool]:
        """Return (Excel row, created)."""
        ...


class Win32ComOutpatientScheduleBackend:
    """Write one authoritative outpatient recurring row on a copied workbook only."""

    @staticmethod
    def _ensure_sheet(workbook):
        try:
            ws = workbook.Worksheets(OUTPATIENT_SCHEDULE_SHEET)
        except Exception:
            ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
            ws.Name = OUTPATIENT_SCHEDULE_SHEET

        for col, header in enumerate(OUTPATIENT_REQUIRED_HEADERS, start=1):
            current = str(ws.Cells(1, col).Value or "").strip()
            if current and current != header:
                raise OutpatientScheduleWriteError(
                    f"{OUTPATIENT_SCHEDULE_SHEET} column {col} must be {header!r}, found {current!r}"
                )
            ws.Cells(1, col).Value = header
        return ws

    @staticmethod
    def _last_value_row(ws) -> int:
        last = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)  # xlUp
        return max(last, 1)

    def upsert(
        self,
        workbook_path: Path,
        request: OutpatientScheduleRequest,
        *,
        patient_name: str,
    ) -> tuple[int, bool]:
        if sys.platform != "win32":
            raise OutpatientScheduleWriteError(
                "Outpatient schedule preview write-back requires Windows with Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise OutpatientScheduleWriteError(
                "pywin32 is required for outpatient schedule preview write-back"
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
                str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False
            )
            ws = self._ensure_sheet(workbook)

            if request.target_base_entry_id:
                row = _target_row_from_entry_id(request.target_base_entry_id)
                current_id = str(ws.Cells(row, 1).Value or "").strip()
                if not current_id:
                    raise OutpatientScheduleWriteError(
                        f"Cannot update {request.target_base_entry_id!r}: row is empty"
                    )
                if current_id.casefold() != request.patient_id.strip().casefold():
                    raise OutpatientScheduleWriteError(
                        f"Cannot update {request.target_base_entry_id!r}: PatientID mismatch"
                    )
                created = False
            else:
                row = self._last_value_row(ws) + 1
                if row < 2:
                    row = 2
                created = True

            values = (
                request.patient_id.strip(),
                patient_name.strip(),
                request.treatment.strip(),
                request.start_time,
                request.day_pattern.strip(),
                (request.therapist_id or "").strip(),
            )
            for col, value in enumerate(values, start=1):
                ws.Cells(row, col).Value = value
            ws.Cells(row, 4).NumberFormat = "hh:mm"
            workbook.Save()
            return row, created
        except OutpatientScheduleWriteError:
            raise
        except Exception as exc:
            raise OutpatientScheduleWriteError(
                f"Excel outpatient schedule write failed: {exc}"
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


def create_outpatient_schedule_preview(
    source_path: str | Path,
    output_path: str | Path,
    request: OutpatientScheduleRequest,
    *,
    backend: OutpatientScheduleBackend | None = None,
    overwrite: bool = False,
) -> OutpatientSchedulePreviewReport:
    """Append/update one outpatient recurring assignment in a NEW .xlsm copy."""

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise OutpatientScheduleWriteError(f"Source workbook not found: {source}")
    if source == output:
        raise OutpatientScheduleWriteError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise OutpatientScheduleWriteError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise OutpatientScheduleWriteError(f"Output already exists: {output}")

    patients = read_patient_registry(source)
    patient = validate_outpatient_schedule_request(request, patients=patients)

    # Validate update identity against authoritative source before copying/writing.
    if request.target_base_entry_id:
        existing = {
            entry.base_entry_id: entry
            for entry in read_outpatient_schedule(source, patients=patients)
        }
        current = existing.get(request.target_base_entry_id)
        if current is None:
            raise OutpatientScheduleWriteError(
                f"Unknown outpatient base_entry_id {request.target_base_entry_id!r}"
            )
        if current.patient_id != request.patient_id.strip():
            raise OutpatientScheduleWriteError(
                f"Cannot move {request.target_base_entry_id!r} to another patient"
            )

    before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        try:
            output.unlink()
        except PermissionError as exc:
            raise OutpatientScheduleWriteError(
                f"Preview workbook is open or locked: {output}"
            ) from exc
    shutil.copy2(source, output)

    selected = backend or Win32ComOutpatientScheduleBackend()
    try:
        row, created = selected.upsert(
            output,
            request,
            patient_name=patient.display_name,
        )
    except Exception:
        try:
            output.unlink()
        except Exception:
            pass
        raise

    after = _sha256(source)
    if before != after:
        try:
            output.unlink()
        except Exception:
            pass
        raise OutpatientScheduleWriteError("Source workbook changed during preview creation")

    expected_id = f"outpatient:{row}"
    entries = read_outpatient_schedule(output)
    matches = [entry for entry in entries if entry.base_entry_id == expected_id]
    verified = (
        len(matches) == 1
        and matches[0].patient_id == request.patient_id.strip()
        and matches[0].treatment == request.treatment.strip()
        and matches[0].start_time == request.start_time
        and matches[0].day_pattern == request.day_pattern.strip()
        and (matches[0].therapist_id or "") == ((request.therapist_id or "").strip())
    )
    if not verified:
        try:
            output.unlink()
        except Exception:
            pass
        raise OutpatientScheduleWriteError(
            "Preview verification failed: outpatient recurring row was not read back exactly"
        )

    return OutpatientSchedulePreviewReport(
        source_path=str(source),
        output_path=str(output),
        base_entry_id=expected_id,
        excel_row=row,
        source_unchanged=True,
        verified_in_output=True,
        created=created,
    )
