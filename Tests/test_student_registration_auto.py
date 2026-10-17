from datetime import date

from rehab_core.models import Student
from rehab_core.registration import NewStudentRequest
from rehab_core.student_ids import next_sequential_student_id
from rehab_excel.student_registration_auto import resolve_student_id_request


def _student(student_id: str) -> Student:
    return Student(
        student_id=student_id,
        display_name=f"Student {student_id}",
        placement_start=date(2026, 9, 1),
        placement_end=date(2026, 12, 31),
    )


def _request(student_id: str = "") -> NewStudentRequest:
    return NewStudentRequest(
        student_id=student_id,
        display_name="New Student",
        placement_start=date(2027, 1, 1),
        placement_end=date(2027, 3, 31),
    )


def test_next_student_id_starts_at_one_when_registry_has_no_numeric_ids():
    assert next_sequential_student_id([_student("LEGACY")]) == "1"


def test_next_student_id_uses_largest_numeric_value_and_ignores_legacy_ids():
    existing = [_student("2"), _student("10"), _student("STU-99"), _student("7")]
    assert next_sequential_student_id(existing) == "11"


def test_blank_student_id_is_resolved_automatically():
    resolved = resolve_student_id_request(
        _request(),
        existing_students=[_student("4"), _student("8")],
    )
    assert resolved.student_id == "9"
    assert resolved.display_name == "New Student"


def test_explicit_student_id_remains_supported_for_legacy_tooling():
    request = _request("STU-IMPORT")
    resolved = resolve_student_id_request(request, existing_students=[_student("8")])
    assert resolved is request
