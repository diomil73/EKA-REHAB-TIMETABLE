from __future__ import annotations

from typing import Iterable

from .models import Student


def next_sequential_student_id(existing_students: Iterable[Student]) -> str:
    """Return the next numeric StudentID without reusing existing numeric IDs.

    Legacy/non-numeric StudentIDs are preserved and ignored for sequencing.
    Numeric IDs are compared by integer value, so the next ID is always greater
    than every existing numeric StudentID.
    """

    numeric_ids: list[int] = []
    for student in existing_students:
        value = str(student.student_id).strip()
        if value.isdigit():
            numeric_ids.append(int(value))
    return str(max(numeric_ids, default=0) + 1)
