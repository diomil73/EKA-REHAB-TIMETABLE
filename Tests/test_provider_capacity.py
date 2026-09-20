from datetime import date, time

from rehab_core.capacity import validate_new_patient_assignment_capacity
from rehab_core.models import BaseScheduleEntry, Session, Therapist
from rehab_core.replacements import find_replacement_candidates
from rehab_core.workload import calculate_therapist_workload


D = date(2026, 9, 18)
SLOTS = [time(8,30), time(9,15), time(10,0), time(10,45), time(11,30), time(12,15), time(13,0)]


def _session(i: int, therapist: str, slot: time, patient: str | None = None) -> Session:
    return Session(
        session_id=f"s{i}",
        patient_id=patient or f"p{i}",
        therapist_id=therapist,
        session_date=D,
        start_time=slot,
    )


def test_two_patients_in_same_clock_slot_count_as_one_capacity_timeslot():
    sessions = (
        _session(1, "T", time(10, 0), "p1"),
        _session(2, "T", time(10, 0), "p2"),
        _session(3, "T", time(11, 30), "p3"),
    )
    workload = calculate_therapist_workload("T", D, sessions)
    assert workload.active_sessions == 3
    assert workload.active_timeslots == 2


def test_therapist_at_six_timeslots_is_not_a_replacement_candidate_for_seventh():
    target = _session(100, "ORIGINAL", time(13, 0), "patient")
    occupied = tuple(_session(i, "FULL", slot) for i, slot in enumerate(SLOTS[:6], 1))
    candidates = find_replacement_candidates(
        target_session=target,
        therapists=(Therapist("FULL", "Full"),),
        sessions=(target, *occupied),
        timeslots=SLOTS,
    )
    assert candidates == []


def test_therapist_at_five_timeslots_can_take_one_more_and_reports_capacity():
    target = _session(100, "ORIGINAL", time(13, 0), "patient")
    occupied = tuple(_session(i, "T", slot) for i, slot in enumerate(SLOTS[:5], 1))
    candidates = find_replacement_candidates(
        target_session=target,
        therapists=(Therapist("T", "Therapist"),),
        sessions=(target, *occupied),
        timeslots=SLOTS,
    )
    assert len(candidates) == 1
    candidate = candidates[0]
    assert candidate.capacity_limit == 6
    assert candidate.capacity_remaining == 1
    assert time(13, 0) in candidate.available_timeslots


def test_new_patient_recurring_assignment_is_rejected_if_any_pattern_day_would_exceed_six():
    therapist = Therapist("T", "Therapist")
    # Six distinct Wednesday timeslots already exist. Monday only has five.
    entries = []
    for i, slot in enumerate(SLOTS[:6]):
        entries.append(
            BaseScheduleEntry(
                base_entry_id=f"wed-{i}",
                patient_id=f"pw{i}",
                treatment="ΦΘ",
                start_time=slot,
                day_pattern="Τε",
                therapist_id="T",
            )
        )
    for i, slot in enumerate(SLOTS[:5]):
        entries.append(
            BaseScheduleEntry(
                base_entry_id=f"mon-{i}",
                patient_id=f"pm{i}",
                treatment="ΦΘ",
                start_time=slot,
                day_pattern="Δε",
                therapist_id="T",
            )
        )

    check = validate_new_patient_assignment_capacity(
        therapist=therapist,
        existing_entries=entries,
        day_pattern="Δε-Τε",
        start_time=time(13, 0),
    )
    assert not check.allowed
    assert len(check.issues) == 1
    issue = check.issues[0]
    assert issue.used_timeslots_before == 6
    assert issue.used_timeslots_after == 7
    assert issue.max_timeslots == 6
