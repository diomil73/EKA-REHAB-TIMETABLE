from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata
from typing import Iterable, Sequence

from openpyxl import load_workbook

from rehab_core.models import Patient, PatientType

from .patient_registry_source import read_patient_registry
from .reader import read_settings


class MasterProjectionError(RuntimeError):
    """Raised when MASTER_SCHEDULE cannot be projected safely."""


@dataclass(frozen=True)
class MasterProjectionRow:
    target_row: int
    planner_row: int
    patient_id: str
    room: str
    display_name: str
    infectious: bool
    status: str | None
    responsible_doctor: str | None = None


def _sort_text(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().casefold())
    return "".join(char for char in decomposed if not unicodedata.combining(char))


def _clinic_key(room: str | None) -> str:
    text = str(room or "").strip().upper()
    if not text:
        return ""
    return text[0]


def _patient_rows(path: str | Path) -> dict[str, int]:
    workbook_path = Path(path)
    wb = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
        keep_vba=workbook_path.suffix.casefold() == ".xlsm",
    )
    try:
        if "PATIENTS" not in wb.sheetnames:
            raise MasterProjectionError("Workbook has no PATIENTS sheet")
        ws = wb["PATIENTS"]
        result: dict[str, int] = {}
        for row in range(2, ws.max_row + 1):
            raw = ws.cell(row, 1).value
            if raw is None:
                continue
            if isinstance(raw, float) and raw.is_integer():
                patient_id = str(int(raw))
            else:
                patient_id = str(raw).strip()
            if not patient_id:
                continue
            if patient_id in result:
                raise MasterProjectionError(
                    f"Duplicate PatientID {patient_id!r} in PATIENTS"
                )
            result[patient_id] = row
        return result
    finally:
        wb.close()


def _master_header_row(path: str | Path) -> int:
    workbook_path = Path(path)
    wb = load_workbook(
        workbook_path,
        read_only=True,
        data_only=False,
        keep_vba=workbook_path.suffix.casefold() == ".xlsm",
    )
    try:
        if "MASTER_SCHEDULE" not in wb.sheetnames:
            return 1
        ws = wb["MASTER_SCHEDULE"]
        for row in range(1, min(ws.max_row, 10) + 1):
            room = str(ws.cell(row, 2).value or "").strip().casefold()
            patient = str(ws.cell(row, 3).value or "").strip().casefold()
            if room == "θάλαμος".casefold() and patient == "ασθενής".casefold():
                return row
    finally:
        wb.close()
    return 1


def build_master_projection(
    workbook_path: str | Path,
    *,
    patients: Iterable[Patient] | None = None,
    room_order: Sequence[str] | None = None,
) -> tuple[MasterProjectionRow, ...]:
    path = Path(workbook_path)
    patient_list = tuple(patients) if patients is not None else tuple(read_patient_registry(path))
    rooms = tuple(room_order) if room_order is not None else read_settings(path).rooms
    row_by_id = _patient_rows(path)

    room_rank = {_sort_text(room): index for index, room in enumerate(rooms)}

    inpatients = [
        patient
        for patient in patient_list
        if patient.patient_type == PatientType.INPATIENT
    ]

    def sort_key(patient: Patient):
        room = patient.room or ""
        normalized_room = _sort_text(room)
        return (
            room_rank.get(normalized_room, len(room_rank)),
            normalized_room,
            _sort_text(patient.display_name),
            _sort_text(patient.patient_id),
        )

    ordered = sorted(inpatients, key=sort_key)
    projection: list[MasterProjectionRow] = []
    next_target_row = _master_header_row(path) + 1
    previous_clinic = ""

    for patient in ordered:
        clinic = _clinic_key(patient.room)
        if previous_clinic and clinic and clinic != previous_clinic:
            # Reserve exactly one visual separator row between clinics.
            next_target_row += 1

        planner_row = row_by_id.get(patient.patient_id)
        if planner_row is None:
            raise MasterProjectionError(
                f"PatientID {patient.patient_id!r} has no PATIENTS row"
            )
        projection.append(
            MasterProjectionRow(
                target_row=next_target_row,
                planner_row=planner_row,
                patient_id=patient.patient_id,
                room=patient.room or "",
                display_name=patient.display_name,
                infectious=patient.infectious,
                status=patient.status,
                responsible_doctor=patient.responsible_doctor,
            )
        )
        next_target_row += 1
        if clinic:
            previous_clinic = clinic

    return tuple(projection)


_PLANNER_REF_RE = re.compile(r"(PATIENT_PLANNER!\$?[A-Z]+\$?)(\d+)", re.IGNORECASE)


def remap_master_formula(formula: str, planner_row: int) -> str:
    text = str(formula or "")
    if not text.startswith("="):
        raise MasterProjectionError(f"MASTER template cell is not a formula: {text!r}")
    if planner_row < 2:
        raise MasterProjectionError(f"Invalid PATIENT_PLANNER row: {planner_row}")

    replaced, count = _PLANNER_REF_RE.subn(
        lambda match: f"{match.group(1)}{planner_row}",
        text,
    )
    if count < 1:
        raise MasterProjectionError(
            f"MASTER formula has no PATIENT_PLANNER reference: {text!r}"
        )
    return replaced
