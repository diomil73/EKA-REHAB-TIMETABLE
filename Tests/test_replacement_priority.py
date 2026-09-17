from datetime import date, time

from rehab_core import Session, Therapist, find_replacement_candidates


DAY = date(2026, 9, 17)
TARGET_TIME = time(12, 15)
ALT_TIME = time(13, 0)


def _target() -> Session:
    return Session("TARGET", "P-TARGET", "T-ORIGINAL", DAY, TARGET_TIME)


def test_busy_exact_time_does_not_exclude_provider_when_alternative_slot_exists():
    sessions = [
        Session("A-BUSY", "P-A", "T-A", DAY, TARGET_TIME),
    ]

    candidates = find_replacement_candidates(
        _target(),
        therapists=[Therapist("T-A", "A")],
        sessions=sessions,
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-A"]
    assert candidates[0].exact_time_available is False
    assert candidates[0].available_timeslots == (ALT_TIME,)


def test_lower_load_ranks_before_exact_time_match():
    sessions = [
        # A has load 1 and is busy at the requested time, but free at 13:00.
        Session("A1", "P-A1", "T-A", DAY, TARGET_TIME),
        # B has load 2 but is free at the requested time.
        Session("B1", "P-B1", "T-B", DAY, time(9, 0)),
        Session("B2", "P-B2", "T-B", DAY, time(10, 0)),
    ]

    candidates = find_replacement_candidates(
        _target(),
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B")],
        sessions=sessions,
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-A", "T-B"]
    assert candidates[0].active_sessions == 1
    assert candidates[0].exact_time_available is False
    assert candidates[1].active_sessions == 2
    assert candidates[1].exact_time_available is True


def test_exact_time_match_breaks_tie_when_load_is_equal():
    sessions = [
        Session("A1", "P-A1", "T-A", DAY, TARGET_TIME),
        Session("B1", "P-B1", "T-B", DAY, time(9, 0)),
    ]

    candidates = find_replacement_candidates(
        _target(),
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B")],
        sessions=sessions,
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-B", "T-A"]
    assert candidates[0].active_sessions == candidates[1].active_sessions == 1
    assert candidates[0].exact_time_available is True
    assert candidates[1].exact_time_available is False


def test_provider_without_any_common_free_timeslot_is_excluded():
    sessions = [
        Session("A1", "P-A1", "T-A", DAY, TARGET_TIME),
        Session("A2", "P-A2", "T-A", DAY, ALT_TIME),
    ]

    candidates = find_replacement_candidates(
        _target(),
        therapists=[Therapist("T-A", "A")],
        sessions=sessions,
        timeslots=[TARGET_TIME, ALT_TIME],
    )

    assert candidates == []
