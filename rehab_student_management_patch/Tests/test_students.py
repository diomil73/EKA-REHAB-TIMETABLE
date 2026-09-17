from datetime import date, time

from rehab_core import (
    Session,
    Student,
    StudentAssignment,
    Therapist,
    build_student_display_map,
    find_replacement_candidates,
    student_display_state,
    students_for_session,
)


def _student() -> Student:
    return Student(
        student_id="STU-001",
        display_name="Μαρία Παπαδοπούλου",
        student_number=1,
        placement_start=date(2026, 9, 1),
        placement_end=date(2026, 12, 20),
        supervisor_therapist_id="TH-001",
    )


def test_active_student_shows_real_name_and_green_font():
    state = student_display_state(_student(), date(2026, 10, 15))

    assert state.label == "Μαρία Παπαδοπούλου"
    assert state.active is True
    assert state.use_green_font is True


def test_finished_student_is_anonymized_but_identity_is_retained():
    student = _student()
    state = student_display_state(student, date(2027, 1, 10))

    assert state.label == "φοιτητής 1"
    assert state.active is False
    assert state.use_green_font is False
    assert student.display_name == "Μαρία Παπαδοπούλου"


def test_student_number_produces_stable_generic_label():
    student = Student(
        student_id="STU-010",
        display_name="Νίκος Δοκιμή",
        student_number=7,
        placement_start=date(2026, 2, 1),
        placement_end=date(2026, 6, 30),
    )

    state = student_display_state(student, date(2026, 9, 1))
    assert state.label == "φοιτητής 7"


def test_build_display_map_handles_active_and_finished_students():
    active = _student()
    finished = Student(
        student_id="STU-002",
        display_name="Γιώργος Δοκιμή",
        student_number=2,
        placement_start=date(2026, 2, 1),
        placement_end=date(2026, 6, 30),
    )

    states = build_student_display_map([finished, active], date(2026, 10, 1))

    assert states["STU-001"].label == "Μαρία Παπαδοπούλου"
    assert states["STU-002"].label == "φοιτητής 2"


def test_students_are_attached_to_session_without_replacing_therapist():
    student = _student()
    assignment = StudentAssignment(
        assignment_id="SA-1",
        student_id=student.student_id,
        session_id="S-1",
    )

    result = students_for_session("S-1", [assignment], [student])

    assert result == [student]


def test_student_is_not_a_replacement_candidate():
    target = Session(
        session_id="S-1",
        patient_id="P-1",
        therapist_id="TH-001",
        session_date=date(2026, 10, 1),
        start_time=time(12, 15),
    )
    therapist = Therapist("TH-002", "Θεραπευτής Β")
    student = _student()

    candidates = find_replacement_candidates(
        target_session=target,
        therapists=[therapist],
        sessions=[target],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["TH-002"]
    assert all(candidate.therapist_id != student.student_id for candidate in candidates)
