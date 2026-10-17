from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata
from typing import Iterable, Sequence

from rehab_core.models import Patient, PatientType

from .package_probe import WorkbookPackageProbe
from .patient_registry_source import read_patient_registry


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
    text = str(room or "").strip().casefold()
    if not text:
        return ""
    if text[0] in {"a", "α"}:
        return "A"
    if text[0] in {"b", "β"}:
        return "B"
    return ""


def _room_number(room: str | None) -> int | None:
    text = str(room or "").strip()
    match = re.search(r"(\d+)", text)
    if match is None:
        return None
    return int(match.group(1))


def _package_patient_id(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        numeric = float(text)
    except ValueError:
        return text
    if numeric.is_integer():
        return str(int(numeric))
    return text


def _projection_metadata(
    path: str | Path,
    *,
    room_order: Sequence[str] | None,
) -> tuple[tuple[str, ...], dict[str, int], int]:
    """Read lightweight MASTER metadata directly from the OOXML package.

    These values are stored workbook data and do not require Excel formula
    evaluation. Reading the package directly avoids loading the full macro-enabled
    workbook with openpyxl during every daily restore.
    """

    probe = WorkbookPackageProbe(path)
    sheet_names = set(probe.sheet_names)
    if "PATIENTS" not in sheet_names:
        raise MasterProjectionError("Workbook has no PATIENTS sheet")

    patient_values = probe.cell_values("PATIENTS")
    row_by_id: dict[str, int] = {}
    for address, raw in patient_values.items():
        match = re.fullmatch(r"A(\d+)", address)
        if match is None:
            continue
        row = int(match.group(1))
        if row < 2:
            continue
        patient_id = _package_patient_id(raw)
        if not patient_id:
            continue
        if patient_id in row_by_id:
            raise MasterProjectionError(
                f"Duplicate PatientID {patient_id!r} in PATIENTS"
            )
        row_by_id[patient_id] = row

    if room_order is None:
        if "SETTINGS" not in sheet_names:
            raise MasterProjectionError("Workbook has no SETTINGS sheet")
        settings_values = probe.cell_values("SETTINGS")
        room_rows: list[tuple[int, str]] = []
        for address, raw in settings_values.items():
            match = re.fullmatch(r"H(\d+)", address)
            if match is None:
                continue
            row = int(match.group(1))
            if row < 2:
                continue
            value = str(raw or "").strip()
            if value:
                room_rows.append((row, value))
        rooms = tuple(value for _row, value in sorted(room_rows))
    else:
        rooms = tuple(room_order)

    header_row = 1
    if "MASTER_SCHEDULE" in sheet_names:
        master_values = probe.cell_values("MASTER_SCHEDULE")
        for row in range(1, 11):
            room = str(master_values.get(f"B{row}", "") or "").strip().casefold()
            patient = str(master_values.get(f"C{row}", "") or "").strip().casefold()
            if room == "θάλαμος".casefold() and patient == "ασθενής".casefold():
                header_row = row
                break

    return rooms, row_by_id, header_row


def build_master_projection(
    workbook_path: str | Path,
    *,
    patients: Iterable[Patient] | None = None,
    room_order: Sequence[str] | None = None,
) -> tuple[MasterProjectionRow, ...]:
    path = Path(workbook_path)
    patient_list = tuple(patients) if patients is not None else tuple(read_patient_registry(path))
    rooms, row_by_id, master_header_row = _projection_metadata(
        path,
        room_order=room_order,
    )

    room_rank = {_sort_text(room): index for index, room in enumerate(rooms)}

    inpatients = [
        patient
        for patient in patient_list
        if patient.patient_type == PatientType.INPATIENT
    ]

    def sort_key(patient: Patient):
        room = patient.room or ""
        normalized_room = _sort_text(room)
        clinic = _clinic_key(room)
        number = _room_number(room)

        if clinic in {"A", "B"} and number is not None:
            return (
                0 if clinic == "A" else 1,
                number,
                normalized_room,
                _sort_text(patient.display_name),
                _sort_text(patient.patient_id),
            )

        return (
            2,
            room_rank.get(normalized_room, len(room_rank)),
            normalized_room,
            _sort_text(patient.display_name),
            _sort_text(patient.patient_id),
        )

    ordered = sorted(inpatients, key=sort_key)
    projection: list[MasterProjectionRow] = []
    next_target_row = master_header_row + 1
    previous_clinic = ""

    for patient in ordered:
        clinic = _clinic_key(patient.room)
        if clinic and clinic != previous_clinic:
            next_target_row += 1
            previous_clinic = clinic

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
