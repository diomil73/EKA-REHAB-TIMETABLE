from datetime import date, time
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_core.models import AbsenceKind, DailyAbsence, Patient, Session, Therapist
from rehab_core.replacements import find_replacement_candidates


TODAY = date(2026, 9, 17)
SLOT = time(12, 15)


def target_session(*, robotic: bool = False, patient_id: str = "P-TARGET") -> Session:
    return Session(
        session_id="TARGET",
        patient_id=patient_id,
        therapist_id="T-ORIGINAL",
        session_date=TODAY,
        start_time=SLOT,
        robotic=robotic,
    )


def test_original_therapist_is_never_suggested():
    candidates = find_replacement_candidates(
        target_session(),
        therapists=[
            Therapist("T-ORIGINAL", "Original"),
            Therapist("T-FREE", "Free"),
        ],
        sessions=[],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-FREE"]


def test_busy_therapist_is_excluded():
    busy_session = Session(
        session_id="BUSY",
        patient_id="P-BUSY",
        therapist_id="T-BUSY",
        session_date=TODAY,
        start_time=SLOT,
    )

    candidates = find_replacement_candidates(
        target_session(),
        therapists=[Therapist("T-BUSY", "Busy"), Therapist("T-FREE", "Free")],
        sessions=[busy_session],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-FREE"]


def test_patient_absence_frees_a_candidate_at_same_time():
    busy_session = Session(
        session_id="BUSY",
        patient_id="P-BUSY",
        therapist_id="T-CANDIDATE",
        session_date=TODAY,
        start_time=SLOT,
    )
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P-BUSY",
        absence_date=TODAY,
    )

    candidates = find_replacement_candidates(
        target_session(),
        therapists=[Therapist("T-CANDIDATE", "Candidate")],
        sessions=[busy_session],
        absences=[absence],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-CANDIDATE"]


def test_absent_therapist_is_excluded():
    absence = DailyAbsence(
        absence_kind=AbsenceKind.THERAPIST,
        subject_id="T-ABSENT",
        absence_date=TODAY,
    )

    candidates = find_replacement_candidates(
        target_session(),
        therapists=[Therapist("T-ABSENT", "Absent"), Therapist("T-FREE", "Free")],
        sessions=[],
        absences=[absence],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-FREE"]


def test_lower_daily_workload_is_ranked_first():
    sessions = [
        Session("A1", "P1", "T-A", TODAY, time(9, 0)),
        Session("A2", "P2", "T-A", TODAY, time(10, 0)),
        Session("B1", "P3", "T-B", TODAY, time(9, 0)),
    ]

    candidates = find_replacement_candidates(
        target_session(),
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B")],
        sessions=sessions,
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-B", "T-A"]
    assert candidates[0].active_timeslots == 1
    assert candidates[1].active_timeslots == 2
    assert candidates[0].active_sessions == 1
    assert candidates[1].active_sessions == 2


def test_more_patients_can_rank_ahead_when_they_use_fewer_timeslots():
    sessions = [
        Session("A1", "PA1", "T-A", TODAY, time(8, 30)),
        Session("A2", "PA2", "T-A", TODAY, time(8, 30)),
        Session("A3", "PA3", "T-A", TODAY, time(9, 15)),
        Session("A4", "PA4", "T-A", TODAY, time(10, 0)),
        Session("A5", "PA5", "T-A", TODAY, time(10, 45)),
        Session("A6", "PA6", "T-A", TODAY, time(10, 45)),
        Session("B1", "PB1", "T-B", TODAY, time(8, 30)),
        Session("B2", "PB2", "T-B", TODAY, time(9, 15)),
        Session("B3", "PB3", "T-B", TODAY, time(10, 0)),
        Session("B4", "PB4", "T-B", TODAY, time(10, 45)),
        Session("B5", "PB5", "T-B", TODAY, time(11, 30)),
    ]

    candidates = find_replacement_candidates(
        target_session(),
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B")],
        sessions=sessions,
        timeslots=[time(8, 30), time(9, 15), time(10, 0), time(10, 45), time(11, 30), SLOT],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-A", "T-B"]
    assert candidates[0].active_sessions == 6
    assert candidates[0].active_timeslots == 4
    assert candidates[0].capacity_remaining == 2
    assert candidates[1].active_sessions == 5
    assert candidates[1].active_timeslots == 5
    assert candidates[1].capacity_remaining == 1


def test_patient_absence_is_removed_from_daily_workload():
    sessions = [
        Session("A1", "P1", "T-A", TODAY, time(9, 0)),
        Session("A2", "P2", "T-A", TODAY, time(10, 0)),
        Session("B1", "P3", "T-B", TODAY, time(9, 0)),
    ]
    absences = [
        DailyAbsence(
            absence_kind=AbsenceKind.PATIENT,
            subject_id="P2",
            absence_date=TODAY,
        )
    ]

    candidates = find_replacement_candidates(
        target_session(),
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B")],
        sessions=sessions,
        absences=absences,
    )

    assert candidates[0].active_sessions == 1
    assert candidates[1].active_sessions == 1
    assert candidates[0].active_timeslots == 1
    assert candidates[1].active_timeslots == 1


def test_robotic_session_only_returns_robotic_capable_therapists():
    candidates = find_replacement_candidates(
        target_session(robotic=True),
        therapists=[
            Therapist("T-NORMAL", "Normal", robotic_capable=False),
            Therapist("T-ROBOT", "Robot", robotic_capable=True),
        ],
        sessions=[],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-ROBOT"]


def test_total_load_stays_first_priority_for_infectious_target():
    patients = [
        Patient("P-TARGET", "Target", infectious=True),
        Patient("P-INF", "Infectious", infectious=True),
        Patient("P-NORMAL", "Normal", infectious=False),
    ]
    sessions = [
        Session("A1", "P-INF", "T-A", TODAY, time(9, 0)),
        Session("B1", "P-NORMAL", "T-B", TODAY, time(9, 0)),
        Session("B2", "P-NORMAL", "T-B", TODAY, time(10, 0)),
    ]

    candidates = find_replacement_candidates(
        target_session(patient_id="P-TARGET"),
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B")],
        sessions=sessions,
        patients=patients,
    )

    # Operational timeslot load remains the first priority. Infectious workload
    # is a later tie-breaker and never overrides a lower occupied-timeslot load.
    assert [candidate.therapist_id for candidate in candidates] == ["T-A", "T-B"]
    assert candidates[0].active_timeslots == 1
    assert candidates[1].active_timeslots == 2


def test_infectious_workload_breaks_tie_after_load_and_patient_count_match():
    patients = [
        Patient("P-TARGET", "Target", infectious=True),
        Patient("P-INF", "Infectious", infectious=True),
        Patient("P-NORMAL", "Normal", infectious=False),
    ]
    sessions = [
        Session("A1", "P-INF", "T-A", TODAY, time(9, 0)),
        Session("B1", "P-NORMAL", "T-B", TODAY, time(9, 0)),
    ]

    candidates = find_replacement_candidates(
        target_session(patient_id="P-TARGET"),
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B")],
        sessions=sessions,
        patients=patients,
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-B", "T-A"]
    assert candidates[0].active_timeslots == candidates[1].active_timeslots == 1
    assert candidates[0].active_sessions == candidates[1].active_sessions == 1
    assert candidates[0].infectious_sessions == 0
    assert candidates[1].infectious_sessions == 1
