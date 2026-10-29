from datetime import date, time

import pytest

from rehab_core.models import Student
from rehab_excel.layout_contract import LayoutSafetyError
from rehab_excel.student_daily_layout import (
    resolve_student_daily_cell_from_cells,
    student_daily_slots_from_cells,
)


TARGET_DATE = date(2026, 10, 26)


def _student(student_id: str, name: str) -> Student:
    return Student(
        student_id=student_id,
        display_name=name,
        placement_start=TARGET_DATE,
        placement_end=TARGET_DATE,
    )


def test_discovers_numbered_student_slot_in_bottom_daily_block():
    cells = {"H11": "φοιτ 1"}

    assert student_daily_slots_from_cells(cells) == ((1, 12, 8),)


def test_single_active_student_maps_to_existing_student_slot():
    student = _student("STU-1", "Μαρία")
    cells = {"H11": "φοιτ 1"}

    cell = resolve_student_daily_cell_from_cells(
        cells,
        student.student_id,
        time(8, 30),
        students={student.student_id: student},
        target_date=TARGET_DATE,
    )

    assert cell == "H12"


def test_multiple_students_map_deterministically_by_student_id_to_numbered_slots():
    first = _student("A-STUDENT", "Άλφα")
    second = _student("B-STUDENT", "Βήτα")
    cells = {"H11": "φοιτ 1", "I11": "φοιτ 2"}
    students = {second.student_id: second, first.student_id: first}

    assert resolve_student_daily_cell_from_cells(
        cells,
        first.student_id,
        time(10, 0),
        students=students,
        target_date=TARGET_DATE,
    ) == "H14"
    assert resolve_student_daily_cell_from_cells(
        cells,
        second.student_id,
        time(10, 0),
        students=students,
        target_date=TARGET_DATE,
    ) == "I14"


def test_student_slot_capacity_overflow_safety_stops_instead_of_colliding():
    first = _student("A-STUDENT", "Άλφα")
    second = _student("B-STUDENT", "Βήτα")

    with pytest.raises(LayoutSafetyError, match="student-slot capacity exceeded"):
        resolve_student_daily_cell_from_cells(
            {"H11": "φοιτ 1"},
            first.student_id,
            time(8, 30),
            students={first.student_id: first, second.student_id: second},
            target_date=TARGET_DATE,
        )


def test_inactive_student_cannot_claim_daily_slot():
    inactive = Student(
        student_id="OLD",
        display_name="Παλιός",
        placement_start=date(2026, 9, 1),
        placement_end=date(2026, 9, 30),
    )

    with pytest.raises(LayoutSafetyError, match="not active"):
        resolve_student_daily_cell_from_cells(
            {"H11": "φοιτ 1"},
            inactive.student_id,
            time(8, 30),
            students={inactive.student_id: inactive},
            target_date=TARGET_DATE,
        )
