from datetime import date, time

from rehab_core.models import Patient, ReplacementAssignment, Session, Therapist
from rehab_core.replacement_options import find_replacement_options

DAY = date(2026, 9, 18)


def test_accepted_replacement_updates_next_provider_load_and_timeslots():
    sessions = [
        Session("S1", "P1", "ABSENT", DAY, time(9, 15), "ΦΘ"),
        Session("S2", "P2", "ABSENT", DAY, time(10, 0), "ΦΘ"),
    ]
    patients = [Patient("P1", "P1"), Patient("P2", "P2")]
    therapists = [Therapist("T1", "T1"), Therapist("T2", "T2")]
    slots = [time(9, 15), time(10, 0), time(10, 45)]

    before = find_replacement_options(
        sessions[0], therapists, sessions, patients=patients, timeslots=slots
    )
    assert before[0].provider_id == "T1"
    assert before[0].active_sessions == 0
    assert time(9, 15) in before[0].available_timeslots

    accepted = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="ABSENT",
        replacement_therapist_id="T1",
        replacement_date=DAY,
        replacement_time=time(9, 15),
    )
    after = find_replacement_options(
        sessions[1],
        therapists,
        sessions,
        patients=patients,
        replacements=[accepted],
        timeslots=slots,
    )
    t1 = next(option for option in after if option.provider_id == "T1")
    assert t1.active_sessions == 1
    assert time(9, 15) not in t1.available_timeslots


def test_sequential_assignment_can_change_next_ranking():
    sessions = [
        Session("S1", "P1", "ABSENT", DAY, time(9, 15), "ΦΘ"),
        Session("S2", "P2", "ABSENT", DAY, time(10, 0), "ΦΘ"),
        Session("BASE", "P3", "T2", DAY, time(13, 0), "ΦΘ"),
    ]
    patients = [Patient("P1", "P1"), Patient("P2", "P2"), Patient("P3", "P3")]
    therapists = [Therapist("T1", "T1"), Therapist("T2", "T2")]
    slots = [time(9, 15), time(10, 0), time(13, 0)]

    first = find_replacement_options(
        sessions[0], therapists, sessions, patients=patients, timeslots=slots
    )
    assert [option.provider_id for option in first[:2]] == ["T1", "T2"]

    accepted = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="ABSENT",
        replacement_therapist_id="T1",
        replacement_date=DAY,
        replacement_time=time(9, 15),
    )
    second = find_replacement_options(
        sessions[1],
        therapists,
        sessions,
        patients=patients,
        replacements=[accepted],
        timeslots=slots,
    )
    assert second[0].active_sessions == 1
    assert second[1].active_sessions == 1
