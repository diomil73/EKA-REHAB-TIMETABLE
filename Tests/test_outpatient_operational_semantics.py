from datetime import date, time

from rehab_core.capacity import therapist_daily_capacity
from rehab_core.models import Patient, PatientType, Session, Therapist
from rehab_core.replacements import find_replacement_candidates
from rehab_core.workload import calculate_therapist_workload


DAY = date(2026, 9, 17)


def test_outpatient_counts_in_therapist_workload_like_inpatient():
    sessions = [
        Session("S-IN", "P-IN", "T1", DAY, time(9, 0)),
        Session("S-OUT", "P-OUT", "T1", DAY, time(10, 0)),
    ]
    patients = [
        Patient("P-IN", "Inpatient"),
        Patient("P-OUT", "Outpatient", patient_type=PatientType.OUTPATIENT),
    ]

    workload = calculate_therapist_workload(
        "T1", DAY, sessions=sessions, patients=patients
    )

    assert workload.active_sessions == 2
    assert workload.active_timeslots == 2


def test_outpatient_counts_toward_daily_capacity():
    therapist = Therapist("T1", "Therapist", max_daily_timeslots=2)
    sessions = [
        Session("S-IN", "P-IN", "T1", DAY, time(9, 0)),
        Session("S-OUT", "P-OUT", "T1", DAY, time(10, 0)),
    ]
    patients = [
        Patient("P-IN", "Inpatient"),
        Patient("P-OUT", "Outpatient", patient_type=PatientType.OUTPATIENT),
    ]

    capacity = therapist_daily_capacity(
        therapist,
        DAY,
        sessions,
        patients=patients,
    )

    assert capacity.used_timeslots == 2
    assert capacity.at_capacity
    assert capacity.remaining_timeslots == 0


def test_outpatient_session_can_receive_replacement_candidates():
    target = Session("S-OUT", "P-OUT", "T-ORIGINAL", DAY, time(10, 0))
    outpatient = Patient(
        "P-OUT",
        "Outpatient",
        patient_type=PatientType.OUTPATIENT,
    )
    candidate = Therapist("T-NEW", "Replacement Therapist")

    candidates = find_replacement_candidates(
        target_session=target,
        therapists=[candidate],
        sessions=[target],
        patients=[outpatient],
    )

    assert [item.therapist_id for item in candidates] == ["T-NEW"]
    assert candidates[0].exact_time_available


def test_patient_type_does_not_change_operational_counting():
    inpatient = Patient("P1", "Same patient", patient_type=PatientType.INPATIENT)
    outpatient = Patient("P1", "Same patient", patient_type=PatientType.OUTPATIENT)
    session = Session("S1", "P1", "T1", DAY, time(11, 0))

    inpatient_workload = calculate_therapist_workload(
        "T1", DAY, sessions=[session], patients=[inpatient]
    )
    outpatient_workload = calculate_therapist_workload(
        "T1", DAY, sessions=[session], patients=[outpatient]
    )

    assert outpatient_workload == inpatient_workload
