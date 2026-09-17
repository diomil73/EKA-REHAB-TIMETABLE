from datetime import date, time

from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    Session,
)
from rehab_core.workload import calculate_therapist_workload


DAY = date(2026, 9, 17)


def test_counts_active_infectious_and_robotic_sessions():
    sessions = [
        Session("S1", "P-INF", "T1", DAY, time(9, 0), robotic=True),
        Session("S2", "P-NORMAL", "T1", DAY, time(10, 0)),
    ]
    patients = [
        Patient("P-INF", "Infectious", infectious=True),
        Patient("P-NORMAL", "Normal", infectious=False),
    ]

    workload = calculate_therapist_workload(
        "T1", DAY, sessions=sessions, patients=patients
    )

    assert workload.active_sessions == 2
    assert workload.infectious_sessions == 1
    assert workload.robotic_sessions == 1
    assert workload.replacement_sessions == 0


def test_absent_patient_is_removed_from_operational_workload():
    sessions = [
        Session("S1", "P1", "T1", DAY, time(9, 0)),
        Session("S2", "P2", "T1", DAY, time(10, 0)),
    ]
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P2",
        absence_date=DAY,
    )

    workload = calculate_therapist_workload(
        "T1", DAY, sessions=sessions, absences=[absence]
    )

    assert workload.active_sessions == 1


def test_replaced_base_session_is_removed_from_original_therapist():
    session = Session("S1", "P1", "T-ORIGINAL", DAY, time(12, 15))
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="T-ORIGINAL",
        replacement_therapist_id="T-NEW",
        replacement_date=DAY,
        replacement_time=time(12, 15),
    )

    workload = calculate_therapist_workload(
        "T-ORIGINAL",
        DAY,
        sessions=[session],
        replacements=[replacement],
    )

    assert workload.active_sessions == 0
    assert workload.replacement_sessions == 0


def test_replacement_is_counted_for_new_therapist():
    session = Session("S1", "P-INF", "T-ORIGINAL", DAY, time(12, 15), robotic=True)
    patient = Patient("P-INF", "Infectious", infectious=True)
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P-INF",
        original_therapist_id="T-ORIGINAL",
        replacement_therapist_id="T-NEW",
        replacement_date=DAY,
        replacement_time=time(13, 0),
    )

    workload = calculate_therapist_workload(
        "T-NEW",
        DAY,
        sessions=[session],
        patients=[patient],
        replacements=[replacement],
    )

    assert workload.active_sessions == 1
    assert workload.infectious_sessions == 1
    assert workload.robotic_sessions == 1
    assert workload.replacement_sessions == 1


def test_absent_replacement_patient_removes_replacement_from_workload():
    session = Session("S1", "P1", "T-ORIGINAL", DAY, time(12, 15))
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="T-ORIGINAL",
        replacement_therapist_id="T-NEW",
        replacement_date=DAY,
        replacement_time=time(13, 0),
    )
    absence = DailyAbsence(
        absence_kind=AbsenceKind.PATIENT,
        subject_id="P1",
        absence_date=DAY,
        start_time=time(12, 30),
        end_time=time(13, 30),
    )

    workload = calculate_therapist_workload(
        "T-NEW",
        DAY,
        sessions=[session],
        absences=[absence],
        replacements=[replacement],
    )

    assert workload.active_sessions == 0
    assert workload.replacement_sessions == 0
