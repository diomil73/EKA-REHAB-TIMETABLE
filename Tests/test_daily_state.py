from datetime import date, time

from rehab_core import (
    AbsenceKind,
    DailyAbsence,
    DailySessionStatus,
    ReplacementAssignment,
    Session,
    build_daily_session_states,
)

DAY = date(2026, 9, 17)


def base_session(*, session_id="S1", patient_id="P1", therapist_id="T1", at=time(12, 15)):
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=at,
    )


def test_active_session_preserves_original_and_effective_values():
    session = base_session()

    state = build_daily_session_states([session])[0]

    assert state.status == DailySessionStatus.ACTIVE
    assert state.original_therapist_id == "T1"
    assert state.effective_therapist_id == "T1"
    assert state.original_time == time(12, 15)
    assert state.effective_time == time(12, 15)
    assert not state.original_should_be_struck_through
    assert not state.needs_replacement


def test_patient_absence_marks_session_inactive_and_struck_through():
    session = base_session()
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P1",
        absence_date=DAY,
        reason="εκτός κλινικής",
    )

    state = build_daily_session_states([session], [absence])[0]

    assert state.status == DailySessionStatus.PATIENT_ABSENT
    assert state.effective_therapist_id is None
    assert state.effective_time is None
    assert state.original_should_be_struck_through
    assert state.reason == "εκτός κλινικής"


def test_therapist_absence_creates_pending_replacement_need():
    session = base_session()
    absence = DailyAbsence(
        absence_kind=AbsenceKind.THERAPIST,
        subject_id="T1",
        absence_date=DAY,
        reason="άδεια",
    )

    state = build_daily_session_states([session], [absence])[0]

    assert state.status == DailySessionStatus.THERAPIST_ABSENT
    assert state.needs_replacement
    assert state.original_should_be_struck_through
    assert state.effective_therapist_id is None
    assert state.reason == "άδεια"


def test_replacement_keeps_original_and_exposes_effective_assignment():
    session = base_session()
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="T1",
        replacement_therapist_id="T2",
        replacement_date=DAY,
        replacement_time=time(12, 15),
        reason="T1 absent",
    )

    state = build_daily_session_states([session], replacements=[replacement])[0]

    assert state.status == DailySessionStatus.REPLACED
    assert state.original_therapist_id == "T1"
    assert state.effective_therapist_id == "T2"
    assert state.original_time == time(12, 15)
    assert state.effective_time == time(12, 15)
    assert state.replacement_id == "R1"
    assert not state.original_should_be_struck_through


def test_replacement_can_move_to_new_time_without_mutating_base_session():
    session = base_session()
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="T1",
        replacement_therapist_id="T2",
        replacement_date=DAY,
        replacement_time=time(13, 0),
    )

    state = build_daily_session_states([session], replacements=[replacement])[0]

    assert session.therapist_id == "T1"
    assert session.start_time == time(12, 15)
    assert state.original_therapist_id == "T1"
    assert state.original_time == time(12, 15)
    assert state.effective_therapist_id == "T2"
    assert state.effective_time == time(13, 0)


def test_patient_absence_has_priority_over_replacement():
    session = base_session()
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P1",
        absence_date=DAY,
        reason="άρρωστος",
    )
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="T1",
        replacement_therapist_id="T2",
        replacement_date=DAY,
        replacement_time=time(12, 15),
    )

    state = build_daily_session_states([session], [absence], [replacement])[0]

    assert state.status == DailySessionStatus.PATIENT_ABSENT
    assert state.effective_therapist_id is None
    assert state.reason == "άρρωστος"


def test_time_specific_patient_absence_at_replacement_time_cancels_effective_replacement():
    session = base_session(at=time(12, 15))
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="T1",
        replacement_therapist_id="T2",
        replacement_date=DAY,
        replacement_time=time(13, 0),
    )
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P1",
        absence_date=DAY,
        start_time=time(12, 45),
        end_time=time(14, 0),
        reason="ιατρική εξέταση",
    )

    state = build_daily_session_states([session], [absence], [replacement])[0]

    assert state.status == DailySessionStatus.PATIENT_ABSENT
    assert state.replacement_id == "R1"
    assert state.effective_therapist_id is None
    assert state.reason == "ιατρική εξέταση"


def test_target_date_filters_other_days():
    other = Session(
        session_id="S2",
        patient_id="P2",
        therapist_id="T2",
        session_date=date(2026, 9, 18),
        start_time=time(10, 0),
    )

    states = build_daily_session_states([base_session(), other], target_date=DAY)

    assert [state.session_id for state in states] == ["S1"]
