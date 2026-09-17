from datetime import date, time

import pytest

from rehab_core import (
    ReplacementProviderKind,
    Session,
    Student,
    StudentAssignment,
    Therapist,
    create_replacement_assignment,
    find_replacement_candidates,
    student_display_state,
)


DAY = date(2026, 9, 17)
TARGET_TIME = time(12, 15)
ALT_TIME = time(13, 0)


def _student(*, max_slots: int = 5, replacement_capable: bool = True) -> Student:
    return Student(
        student_id="STU-1",
        display_name="Μαρία Παπαδοπούλου",
        student_number=1,
        placement_start=date(2026, 9, 1),
        placement_end=date(2026, 12, 20),
        replacement_capable=replacement_capable,
        max_daily_timeslots=max_slots,
    )


def _target() -> Session:
    return Session("TARGET", "P-TARGET", "T-ORIGINAL", DAY, TARGET_TIME)


def _student_load(count: int):
    sessions = []
    assignments = []
    slots = [time(8, 0), time(9, 0), time(10, 0), time(11, 0), time(14, 0)]
    for index in range(count):
        session = Session(
            f"S{index}",
            f"P{index}",
            "T-SUPERVISOR",
            DAY,
            slots[index],
        )
        sessions.append(session)
        assignments.append(StudentAssignment(f"A{index}", "STU-1", session.session_id))
    return sessions, assignments


def test_active_student_keeps_real_name_and_green_display_rule():
    state = student_display_state(_student(), DAY)
    assert state.label == "Μαρία Παπαδοπούλου"
    assert state.use_green_font is True


def test_student_with_four_of_five_slots_can_be_replacement_candidate():
    sessions, assignments = _student_load(4)
    sessions.append(_target())

    candidates = find_replacement_candidates(
        _target(),
        therapists=[],
        students=[_student()],
        student_assignments=assignments,
        sessions=sessions,
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert [candidate.provider_id for candidate in candidates] == ["STU-1"]
    assert candidates[0].provider_kind == ReplacementProviderKind.STUDENT
    assert candidates[0].active_sessions == 4
    assert candidates[0].capacity_remaining == 1


def test_student_at_five_of_five_slots_is_not_candidate():
    sessions, assignments = _student_load(5)
    sessions.append(_target())

    candidates = find_replacement_candidates(
        _target(),
        therapists=[],
        students=[_student()],
        student_assignments=assignments,
        sessions=sessions,
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert candidates == []


def test_busy_student_exact_time_remains_candidate_with_alternative_slot():
    sessions, assignments = _student_load(3)
    busy = Session("S-BUSY", "P-BUSY", "T-SUPERVISOR", DAY, TARGET_TIME)
    sessions.extend([busy, _target()])
    assignments.append(StudentAssignment("A-BUSY", "STU-1", "S-BUSY"))

    candidates = find_replacement_candidates(
        _target(),
        therapists=[],
        students=[_student()],
        student_assignments=assignments,
        sessions=sessions,
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert [candidate.provider_id for candidate in candidates] == ["STU-1"]
    assert candidates[0].active_sessions == 4
    assert candidates[0].exact_time_available is False
    assert ALT_TIME in candidates[0].available_timeslots


def test_student_replacement_can_fill_fifth_slot():
    sessions, assignments = _student_load(4)
    target = _target()
    sessions.append(target)

    overlay = create_replacement_assignment(
        replacement_id="R-STUDENT",
        target_session=target,
        replacement_therapist_id="STU-1",
        therapists=[],
        students=[_student()],
        student_assignments=assignments,
        sessions=sessions,
        replacement_time=TARGET_TIME,
    )

    assert overlay.replacement_therapist_id == "STU-1"
    assert overlay.replacement_provider_kind == ReplacementProviderKind.STUDENT


def test_student_cannot_receive_sixth_slot():
    sessions, assignments = _student_load(5)
    target = _target()
    sessions.append(target)

    with pytest.raises(ValueError, match="capacity"):
        create_replacement_assignment(
            replacement_id="R-STUDENT",
            target_session=target,
            replacement_therapist_id="STU-1",
            therapists=[],
            students=[_student()],
            student_assignments=assignments,
            sessions=sessions,
            replacement_time=TARGET_TIME,
        )


def test_student_not_marked_replacement_capable_is_excluded():
    candidates = find_replacement_candidates(
        _target(),
        therapists=[Therapist("T-A", "A")],
        students=[_student(replacement_capable=False)],
        sessions=[_target()],
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert all(candidate.provider_kind != ReplacementProviderKind.STUDENT for candidate in candidates)
