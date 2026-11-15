from __future__ import annotations

from datetime import date, time
from pathlib import Path
import re
import unicodedata
from typing import Mapping

from rehab_core.models import Student

from .layout_contract import (
    LayoutSafetyError,
    THERAPIST_DAILY_BLOCKS,
    THERAPIST_DAILY_TIMES,
    cell_address,
)
from .package_probe import WorkbookPackageProbe


_STUDENT_SLOT_RE = re.compile(r"^φοιτ(?:ητης)?\.?\s*(\d+)$", re.IGNORECASE)


def _key(value: object) -> str:
    text = " ".join(str(value or "").strip().casefold().split())
    decomposed = unicodedata.normalize("NFD", text)
    return "".join(char for char in decomposed if not unicodedata.combining(char))


def _student_slot_number(value: object) -> int | None:
    match = _STUDENT_SLOT_RE.fullmatch(_key(value))
    return int(match.group(1)) if match else None


def student_daily_slots_from_cells(
    cells: Mapping[str, object],
) -> tuple[tuple[int, int, int], ...]:
    """Return (slot_number, first_time_row, column) for physical student columns."""
    found: list[tuple[int, int, int]] = []
    seen_numbers: set[int] = set()
    for header_row, first_time_row, _last_time_row in THERAPIST_DAILY_BLOCKS:
        for column in range(2, 11):
            slot_number = _student_slot_number(cells.get(cell_address(header_row, column)))
            if slot_number is None:
                continue
            if slot_number in seen_numbers:
                raise LayoutSafetyError(
                    f"Duplicate THERAPIST_DAILY student slot number: {slot_number}"
                )
            seen_numbers.add(slot_number)
            found.append((slot_number, first_time_row, column))
    return tuple(sorted(found, key=lambda item: item[0]))


def resolve_student_daily_cell_from_cells(
    cells: Mapping[str, object],
    student_id: str,
    slot_time: time,
    *,
    students: Mapping[str, Student],
    target_date: date,
) -> str:
    """Resolve an authoritative student to a numbered legacy student column.

    Active authoritative students are mapped deterministically by StudentID to
    the numbered physical ``φοιτ N`` slots already present in THERAPIST_DAILY.
    The function refuses to guess if the sheet exposes fewer physical slots
    than the number of active students for the requested date.
    """
    slots = student_daily_slots_from_cells(cells)
    if not slots:
        raise LayoutSafetyError("THERAPIST_DAILY has no numbered student slot headers")

    active_students = tuple(
        sorted(
            (student for student in students.values() if student.is_active(target_date)),
            key=lambda student: (_key(student.student_id), _key(student.display_name)),
        )
    )
    if len(active_students) > len(slots):
        raise LayoutSafetyError(
            "THERAPIST_DAILY student-slot capacity exceeded: "
            f"{len(active_students)} active student(s), {len(slots)} physical slot(s)"
        )

    target_index = next(
        (index for index, student in enumerate(active_students) if student.student_id == student_id),
        None,
    )
    if target_index is None:
        raise LayoutSafetyError(
            f"Student {student_id!r} is not active on {target_date.isoformat()}"
        )

    normalized_time = slot_time.replace(second=0, microsecond=0)
    try:
        time_index = THERAPIST_DAILY_TIMES.index(normalized_time)
    except ValueError as exc:
        raise LayoutSafetyError(
            f"Unsupported THERAPIST_DAILY time for student {student_id!r}: "
            f"{normalized_time:%H:%M}"
        ) from exc

    _slot_number, first_time_row, column = slots[target_index]
    return cell_address(first_time_row + time_index, column)


def resolve_student_daily_cell(
    workbook_path: str | Path,
    student_id: str,
    slot_time: time,
    *,
    students: Mapping[str, Student],
    target_date: date,
) -> str:
    probe = WorkbookPackageProbe(workbook_path)
    return resolve_student_daily_cell_from_cells(
        probe.cell_values("THERAPIST_DAILY"),
        student_id,
        slot_time,
        students=students,
        target_date=target_date,
    )
