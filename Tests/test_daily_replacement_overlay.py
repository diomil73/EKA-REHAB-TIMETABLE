from datetime import date, time

import pytest

from rehab_core.availability import is_therapist_available
from rehab_core.models import AbsenceKind, DailyAbsence, Patient, Session, Therapist
from rehab_core.replacements import (
    create_replacement_assignment,
    find_replacement_candidates,
)


DAY = date(2026, 9, 17)


def _target_session() -> Session:
    return Session(
        session_id="S-TARGET",
        patient_id="P1",
        therapist_id="T-ORIGINAL",
        session_date=DAY,
        start_time=time(12, 15),
    )


def _therapists() -> list[Therapist]:
    return [
        Therapist("T-ORIGINAL", "Original"),
        Therapist("T-A", "Therapist A"),
        Therapist("T-B", "Therapist B"),
    ]


def test_accepted_replacement_blocks_same_therapist_same_time():
    target = _target_session()
    therapists = _therapists()

    overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-A",
        therapists=therapists,
        sessions=[target],
    )

    assert not is_therapist_available(
        "T-A",
        DAY,
        time(12, 15),
        sessions=[target],
        replacements=[overlay],
    )


def test_accepted_replacement_removes_therapist_from_next_candidate_list():
    target = _target_session()
    second_target = Session(
        session_id="S2",
        patient_id="P2",
        therapist_id="T-ORIGINAL",
        session_date=DAY,
        start_time=time(12, 15),
    )
    therapists = _therapists()

    overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-A",
        therapists=therapists,
        sessions=[target, second_target],
    )

    candidates = find_replacement_candidates(
        target_session=second_target,
        therapists=therapists,
        sessions=[target, second_target],
        replacements=[overlay],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-B"]


def test_replacement_counts_in_operational_workload():
    target = _target_session()
    next_target = Session(
        session_id="S3",
        patient_id="P2",
        therapist_id="T-ORIGINAL",
        session_date=DAY,
        start_time=time(13, 0),
    )
    therapists = _therapists()
    patients = [Patient("P1", "Patient 1"), Patient("P2", "Patient 2")]

    overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-A",
        therapists=therapists,
        sessions=[target, next_target],
    )

    candidates = find_replacement_candidates(
        target_session=next_target,
        therapists=therapists,
        sessions=[target, next_target],
        patients=patients,
        replacements=[overlay],
    )

    by_id = {candidate.therapist_id: candidate for candidate in candidates}
    assert by_id["T-A"].active_sessions == 1
    assert by_id["T-B"].active_sessions == 0
    assert candidates[0].therapist_id == "T-B"


def test_cannot_assign_same_replacement_therapist_twice_at_same_time():
    target = _target_session()
    second_target = Session(
        session_id="S2",
        patient_id="P2",
        therapist_id="T-ORIGINAL",
        session_date=DAY,
        start_time=time(12, 15),
    )
    therapists = _therapists()

    first_overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-A",
        therapists=therapists,
        sessions=[target, second_target],
    )

    with pytest.raises(ValueError, match="not available"):
        create_replacement_assignment(
            replacement_id="R2",
            target_session=second_target,
            replacement_therapist_id="T-A",
            therapists=therapists,
            sessions=[target, second_target],
            replacements=[first_overlay],
        )


def test_base_session_is_not_modified_by_replacement_overlay():
    target = _target_session()
    therapists = _therapists()

    overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-A",
        therapists=therapists,
        sessions=[target],
    )

    assert target.therapist_id == "T-ORIGINAL"
    assert overlay.original_therapist_id == "T-ORIGINAL"
    assert overlay.replacement_therapist_id == "T-A"


def test_replacement_can_use_a_new_time_when_patient_and_therapist_are_free():
    target = _target_session()
    therapists = _therapists()

    overlay = create_replacement_assignment(
        replacement_id="R-MOVED",
        target_session=target,
        replacement_therapist_id="T-A",
        replacement_time=time(13, 0),
        therapists=therapists,
        sessions=[target],
    )

    assert overlay.replacement_time == time(13, 0)
    assert target.start_time == time(12, 15)


def test_replacement_rejects_new_time_when_patient_has_another_session():
    target = _target_session()
    conflict = Session(
        session_id="S-CONFLICT",
        patient_id="P1",
        therapist_id="T-B",
        session_date=DAY,
        start_time=time(13, 0),
    )
    therapists = _therapists()

    with pytest.raises(ValueError, match="Patient is not available"):
        create_replacement_assignment(
            replacement_id="R-MOVED",
            target_session=target,
            replacement_therapist_id="T-A",
            replacement_time=time(13, 0),
            therapists=therapists,
            sessions=[target, conflict],
        )


def test_replacement_rejects_absent_patient():
    target = _target_session()
    therapists = _therapists()
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P1",
        absence_date=DAY,
    )

    with pytest.raises(ValueError, match="Patient is not available"):
        create_replacement_assignment(
            replacement_id="R1",
            target_session=target,
            replacement_therapist_id="T-A",
            therapists=therapists,
            sessions=[target],
            absences=[absence],
        )


def test_same_target_session_cannot_receive_two_replacement_overlays():
    target = _target_session()
    therapists = _therapists()

    first_overlay = create_replacement_assignment(
        replacement_id="R1",
        target_session=target,
        replacement_therapist_id="T-A",
        therapists=therapists,
        sessions=[target],
    )

    with pytest.raises(ValueError, match="already has a replacement"):
        create_replacement_assignment(
            replacement_id="R2",
            target_session=target,
            replacement_therapist_id="T-B",
            therapists=therapists,
            sessions=[target],
            replacements=[first_overlay],
        )
