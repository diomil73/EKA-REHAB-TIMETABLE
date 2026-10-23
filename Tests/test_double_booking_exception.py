from datetime import date, time
from pathlib import Path
import sys

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_core.double_booking_exception import (
    DOUBLE_BOOKING_WARNING,
    confirm_double_booking_exception,
    propose_double_booking_exception,
)
from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    Session,
    Therapist,
)


DAY = date(2026, 9, 18)


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=time(hh, 0),
        treatment="ΦΘ",
    )


def test_double_booking_requires_session_to_be_unresolved_first():
    target = _session("TARGET", "P1", "ABSENT", 10)
    occupied = _session("BUSY", "P2", "T2", 10)

    with pytest.raises(ValueError, match="left unresolved"):
        propose_double_booking_exception(
            target_session=target,
            unresolved_session_ids=(),
            therapist_id="T2",
            replacement_time=time(10, 0),
            therapists=(Therapist("T2", "T2"),),
            sessions=(target, occupied),
        )


def test_double_booking_proposal_requires_a_real_live_slot_conflict():
    target = _session("TARGET", "P1", "ABSENT", 10)

    with pytest.raises(ValueError, match="normally available"):
        propose_double_booking_exception(
            target_session=target,
            unresolved_session_ids=("TARGET",),
            therapist_id="T2",
            replacement_time=time(10, 0),
            therapists=(Therapist("T2", "T2"),),
            sessions=(target,),
        )


def test_double_booking_cannot_override_therapist_absence():
    target = _session("TARGET", "P1", "ABSENT", 10)
    occupied = _session("BUSY", "P2", "T2", 10)

    with pytest.raises(ValueError, match="absent"):
        propose_double_booking_exception(
            target_session=target,
            unresolved_session_ids=("TARGET",),
            therapist_id="T2",
            replacement_time=time(10, 0),
            therapists=(Therapist("T2", "T2"),),
            sessions=(target, occupied),
            absences=(
                DailyAbsence(
                    absence_kind=AbsenceKind.THERAPIST,
                    subject_id="T2",
                    absence_date=DAY,
                ),
            ),
        )


def test_double_booking_cannot_override_patient_collision():
    target = _session("TARGET", "P1", "ABSENT", 10)
    occupied = _session("BUSY", "P2", "T2", 10)
    patient_other = _session("P1-OTHER", "P1", "T3", 10)

    with pytest.raises(ValueError, match="Patient is not available"):
        propose_double_booking_exception(
            target_session=target,
            unresolved_session_ids=("TARGET",),
            therapist_id="T2",
            replacement_time=time(10, 0),
            therapists=(Therapist("T2", "T2"), Therapist("T3", "T3")),
            sessions=(target, occupied, patient_other),
        )


def test_confirmed_double_booking_creates_auditable_replacement_assignment():
    target = _session("TARGET", "P1", "ABSENT", 10)
    occupied = _session("BUSY", "P2", "T2", 10)

    proposal = propose_double_booking_exception(
        target_session=target,
        unresolved_session_ids=("TARGET",),
        therapist_id="T2",
        replacement_time=time(10, 0),
        therapists=(Therapist("T2", "T2"),),
        sessions=(target, occupied),
    )

    assert proposal.warning == DOUBLE_BOOKING_WARNING

    with pytest.raises(ValueError, match="Explicit confirmation"):
        confirm_double_booking_exception(
            proposal,
            target_session=target,
            replacement_id="DB-1",
            confirmed=False,
        )

    assignment = confirm_double_booking_exception(
        proposal,
        target_session=target,
        replacement_id="DB-1",
        confirmed=True,
    )

    assert assignment.target_session_id == "TARGET"
    assert assignment.replacement_therapist_id == "T2"
    assert assignment.replacement_time == time(10, 0)
    assert assignment.reason == "CONFIRMED DOUBLE BOOKING EXCEPTION"
