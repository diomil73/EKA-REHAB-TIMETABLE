from datetime import date, time

from rehab_core.availability import is_patient_available, is_therapist_available
from rehab_core.models import AbsenceKind, DailyAbsence, Session, Therapist
from rehab_core.replacements import create_replacement_assignment


DAY = date(2026, 9, 17)


def test_patient_is_blocked_by_another_base_session_at_same_time():
    sessions = [
        Session("S1", "P1", "T1", DAY, time(11, 0)),
        Session("S2", "P1", "T2", DAY, time(12, 15)),
    ]

    assert not is_patient_available("P1", DAY, time(12, 15), sessions)


def test_patient_absence_makes_patient_unavailable():
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P1",
        absence_date=DAY,
    )

    assert not is_patient_available(
        "P1",
        DAY,
        time(12, 15),
        sessions=[],
        absences=[absence],
    )


def test_target_session_can_be_ignored_when_validating_same_time_replacement():
    target = Session("S1", "P1", "T-OLD", DAY, time(12, 15))

    assert is_patient_available(
        "P1",
        DAY,
        time(12, 15),
        sessions=[target],
        ignore_session_id="S1",
    )


def test_moved_replacement_frees_old_patient_time_and_blocks_new_time():
    target = Session("S1", "P1", "T-OLD", DAY, time(12, 15))
    therapists = [Therapist("T-OLD", "Old"), Therapist("T-NEW", "New")]

    overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-NEW",
        replacement_time=time(13, 0),
        therapists=therapists,
        sessions=[target],
    )

    assert is_patient_available(
        "P1",
        DAY,
        time(12, 15),
        sessions=[target],
        replacements=[overlay],
    )
    assert not is_patient_available(
        "P1",
        DAY,
        time(13, 0),
        sessions=[target],
        replacements=[overlay],
    )


def test_replacement_overlay_frees_original_therapist_base_slot():
    target = Session("S1", "P1", "T-OLD", DAY, time(12, 15))
    therapists = [Therapist("T-OLD", "Old"), Therapist("T-NEW", "New")]

    overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-NEW",
        therapists=therapists,
        sessions=[target],
    )

    assert is_therapist_available(
        "T-OLD",
        DAY,
        time(12, 15),
        sessions=[target],
        replacements=[overlay],
    )
