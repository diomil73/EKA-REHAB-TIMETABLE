from datetime import date, time

from rehab_core import (
    BaseScheduleEntry,
    Session,
    Student,
    StudentAssignment,
    plan_student_end_transition,
    resolve_student_base_entries,
    student_display_state,
)


def _student() -> Student:
    return Student(
        student_id="STU-1",
        display_name="Μαρία Παπαδοπούλου",
        student_number=1,
        placement_start=date(2026, 9, 1),
        placement_end=date(2026, 9, 20),
    )


def _source() -> BaseScheduleEntry:
    return BaseScheduleEntry(
        base_entry_id="B1",
        patient_id="P1",
        treatment="ΦΘ",
        start_time=time(9, 0),
        day_pattern="Δ-Τε-Πα",
        therapist_id="T-SUP",
    )


def _materialized_source() -> Session:
    return Session(
        session_id="B1@2026-09-18",
        patient_id="P1",
        therapist_id="T-SUP",
        session_date=date(2026, 9, 18),
        start_time=time(9, 0),
        treatment="ΦΘ",
    )


def test_expired_student_uses_stable_short_label_but_keeps_identity():
    state = student_display_state(_student(), date(2026, 9, 21))

    assert state.student_id == "STU-1"
    assert state.label == "Φοιτ.1"
    assert state.active is False
    assert state.use_green_font is False


def test_student_assignments_resolve_back_to_unique_recurring_entries():
    source = _source()
    session = _materialized_source()
    assignments = [
        StudentAssignment("A1", "STU-1", session.session_id),
        StudentAssignment("A2", "STU-1", session.session_id),
        StudentAssignment("A-OTHER", "STU-2", session.session_id),
    ]

    affected = resolve_student_base_entries(
        student_id="STU-1",
        student_assignments=assignments,
        sessions=[session],
        base_entries=[source],
    )

    assert len(affected) == 1
    assert affected[0].source_entry == source
    assert affected[0].patient_id == "P1"
    assert affected[0].assignment_ids == ("A1", "A2")


def test_transition_plan_blocks_cross_specialty_patient_collision():
    source = _source()
    psychology = BaseScheduleEntry(
        base_entry_id="PSY-1",
        patient_id="P1",
        treatment="Ψυχ",
        start_time=time(10, 0),
        day_pattern="Δ-Τε-Πα",
        therapist_id="PSYCHOLOGIST-1",
    )
    session = _materialized_source()

    plan = plan_student_end_transition(
        student=_student(),
        student_assignments=[StudentAssignment("A1", "STU-1", session.session_id)],
        sessions=[session],
        base_entries=[source, psychology],
        destination_provider_ids=["T-DEST"],
        standard_timeslots=[time(9, 0), time(10, 0)],
    )

    assert plan.expired_label == "Φοιτ.1"
    assert plan.effective_from == date(2026, 9, 21)
    assert plan.affected_patient_ids == ("P1",)
    assert [(option.destination_provider_id, option.destination_time) for option in plan.options] == [
        ("T-DEST", time(9, 0))
    ]


def test_transition_plan_respects_destination_weekly_capacity():
    source = _source()
    session = _materialized_source()
    occupied_times = [
        time(8, 0),
        time(8, 30),
        time(9, 30),
        time(10, 30),
        time(11, 30),
        time(12, 30),
    ]
    occupied = [
        BaseScheduleEntry(
            base_entry_id=f"D{index}",
            patient_id=f"PX{index}",
            treatment="ΦΘ",
            start_time=slot,
            day_pattern="Δ-Τε-Πα",
            therapist_id="T-FULL",
        )
        for index, slot in enumerate(occupied_times)
    ]

    plan = plan_student_end_transition(
        student=_student(),
        student_assignments=[StudentAssignment("A1", "STU-1", session.session_id)],
        sessions=[session],
        base_entries=[source, *occupied],
        destination_provider_ids=["T-FULL"],
        standard_timeslots=[time(9, 0)],
    )

    assert len(plan.affected_entries) == 1
    assert plan.options == ()


def test_current_supervising_therapist_can_be_suggested_as_takeover_without_mutation():
    source = _source()
    session = _materialized_source()
    assignments = (StudentAssignment("A1", "STU-1", session.session_id),)
    entries = (source,)

    plan = plan_student_end_transition(
        student=_student(),
        student_assignments=assignments,
        sessions=[session],
        base_entries=entries,
        destination_provider_ids=["T-SUP"],
        standard_timeslots=[time(9, 0)],
    )

    assert len(plan.options) == 1
    assert plan.options[0].destination_provider_id == "T-SUP"
    assert plan.options[0].same_time is True
    assert entries == (source,)
    assert assignments == (StudentAssignment("A1", "STU-1", session.session_id),)
