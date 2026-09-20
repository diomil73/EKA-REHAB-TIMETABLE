from datetime import date, time

from rehab_core.capacity_rebalance import build_capacity_rebalance_cases
from rehab_core.models import Session, Therapist

D = date(2026, 9, 18)


def s(sid, patient, provider, hh, mm=0):
    return Session(sid, patient, provider, D, time(hh, mm))


def test_seven_distinct_slots_creates_one_excess_slot_case():
    therapist = Therapist("t", "T")
    sessions = [
        s(f"s{i}", f"p{i}", "t", hh, mm)
        for i, (hh, mm) in enumerate(
            [(8, 30), (9, 15), (10, 0), (10, 45), (11, 30), (12, 15), (13, 0)],
            start=1,
        )
    ]
    cases = build_capacity_rebalance_cases(
        target_date=D, sessions=sessions, therapists=[therapist]
    )
    assert len(cases) == 1
    assert cases[0].used_timeslots == 7
    assert cases[0].max_timeslots == 6
    assert cases[0].excess_timeslots == 1
    assert len(cases[0].slots) == 7


def test_two_patients_same_clock_time_count_as_one_slot_and_stay_grouped():
    therapist = Therapist("t", "T")
    sessions = [
        s("a", "p1", "t", 8, 30),
        s("b", "p2", "t", 8, 30),
        s("c", "p3", "t", 9, 15),
        s("d", "p4", "t", 10, 0),
        s("e", "p5", "t", 10, 45),
        s("f", "p6", "t", 11, 30),
        s("g", "p7", "t", 12, 15),
        s("h", "p8", "t", 13, 0),
    ]
    cases = build_capacity_rebalance_cases(
        target_date=D, sessions=sessions, therapists=[therapist]
    )
    assert len(cases) == 1
    assert cases[0].used_timeslots == 7
    first = next(slot for slot in cases[0].slots if slot.start_time == time(8, 30))
    assert {session.patient_id for session in first.sessions} == {"p1", "p2"}


def test_at_capacity_is_not_a_rebalance_case():
    therapist = Therapist("t", "T")
    sessions = [
        s(str(i), f"p{i}", "t", hh, mm)
        for i, (hh, mm) in enumerate(
            [(8, 30), (9, 15), (10, 0), (10, 45), (11, 30), (12, 15)],
            start=1,
        )
    ]
    cases = build_capacity_rebalance_cases(
        target_date=D, sessions=sessions, therapists=[therapist]
    )
    assert cases == ()
