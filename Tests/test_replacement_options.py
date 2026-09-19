from datetime import date, time

from rehab_core.models import Patient, Session, Therapist
from rehab_core.replacement_options import (
    find_replacement_options,
    order_timeslots_by_preference,
)

D = date(2026, 9, 18)


def session(sid, patient, therapist, hh, mm=0):
    return Session(sid, patient, therapist, D, time(hh, mm))


def test_orders_exact_first_then_nearest_then_chronological():
    result = order_timeslots_by_preference(
        [time(13, 0), time(10, 0), time(8, 30), time(9, 15)],
        time(8, 30),
    )
    assert result == (time(8, 30), time(9, 15), time(10, 0), time(13, 0))


def test_busy_exact_time_provider_remains_with_alternative_slot():
    target = session("target", "p", "orig", 8, 30)
    sessions = [
        target,
        session("busy", "other", "low", 8, 30),
        session("h1", "x1", "high", 9, 15),
        session("h2", "x2", "high", 10, 0),
    ]
    patients = [Patient("p", "P"), Patient("other", "O"), Patient("x1", "X1"), Patient("x2", "X2")]
    therapists = [Therapist("low", "Low"), Therapist("high", "High")]
    options = find_replacement_options(
        target,
        therapists,
        sessions,
        patients=patients,
        timeslots=[time(8, 30), time(9, 15), time(10, 0)],
    )
    low = next(o for o in options if o.provider_id == "low")
    assert low.exact_time_available is False
    assert low.recommended_time == time(9, 15)


def test_lower_load_stays_ahead_even_when_higher_load_has_exact_time():
    target = session("target", "p", "orig", 8, 30)
    sessions = [
        target,
        session("lowbusy", "o", "low", 8, 30),
        session("h1", "x1", "high", 9, 15),
        session("h2", "x2", "high", 10, 0),
    ]
    patients = [Patient("p", "P"), Patient("o", "O"), Patient("x1", "X1"), Patient("x2", "X2")]
    therapists = [Therapist("low", "Low"), Therapist("high", "High")]
    options = find_replacement_options(
        target,
        therapists,
        sessions,
        patients=patients,
        timeslots=[time(8, 30), time(9, 15), time(10, 0), time(10, 45)],
    )
    assert options[0].provider_id == "low"
    assert options[0].active_sessions < options[1].active_sessions
