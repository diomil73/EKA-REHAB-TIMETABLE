from datetime import date, time
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_core.availability import is_therapist_available
from rehab_core.models import AbsenceKind, DailyAbsence, Session


TODAY = date(2026, 9, 17)
SLOT = time(12, 15)


def base_session() -> Session:
    return Session(
        session_id="S-001",
        patient_id="P-001",
        therapist_id="T-001",
        session_date=TODAY,
        start_time=SLOT,
    )


def test_scheduled_patient_blocks_therapist_slot():
    assert not is_therapist_available(
        therapist_id="T-001",
        target_date=TODAY,
        target_time=SLOT,
        sessions=[base_session()],
    )


def test_patient_absence_frees_therapist_slot():
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P-001",
        absence_date=TODAY,
    )

    assert is_therapist_available(
        therapist_id="T-001",
        target_date=TODAY,
        target_time=SLOT,
        sessions=[base_session()],
        absences=[absence],
    )


def test_therapist_absence_keeps_slot_unavailable_even_if_patient_absent():
    absences = [
        DailyAbsence(
            absence_kind=AbsenceKind.PATIENT,
            subject_id="P-001",
            absence_date=TODAY,
        ),
        DailyAbsence(
            absence_kind=AbsenceKind.THERAPIST,
            subject_id="T-001",
            absence_date=TODAY,
        ),
    ]

    assert not is_therapist_available(
        therapist_id="T-001",
        target_date=TODAY,
        target_time=SLOT,
        sessions=[base_session()],
        absences=absences,
    )


def test_partial_patient_absence_only_frees_covered_time():
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P-001",
        absence_date=TODAY,
        start_time=time(12, 0),
        end_time=time(13, 0),
    )

    assert is_therapist_available(
        therapist_id="T-001",
        target_date=TODAY,
        target_time=SLOT,
        sessions=[base_session()],
        absences=[absence],
    )
