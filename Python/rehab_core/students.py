from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable

from .models import Student, StudentAssignment


@dataclass(frozen=True)
class StudentDisplayState:
    student_id: str
    label: str
    active: bool
    use_green_font: bool


def student_display_state(student: Student, target_date: date) -> StudentDisplayState:
    active = student.is_active(target_date)
    label = student.display_name if active else f"Φοιτ.{student.student_number}"
    return StudentDisplayState(
        student_id=student.student_id,
        label=label,
        active=active,
        use_green_font=active,
    )


def build_student_display_map(
    students: Iterable[Student], target_date: date
) -> dict[str, StudentDisplayState]:
    states = [student_display_state(student, target_date) for student in students]
    return {state.student_id: state for state in states}


def students_for_session(
    session_id: str,
    assignments: Iterable[StudentAssignment],
    students: Iterable[Student],
) -> list[Student]:
    student_by_id = {student.student_id: student for student in students}
    result: list[Student] = []
    for assignment in assignments:
        if assignment.session_id != session_id:
            continue
        student = student_by_id.get(assignment.student_id)
        if student is not None:
            result.append(student)
    result.sort(key=lambda student: (student.student_number, student.student_id))
    return result
