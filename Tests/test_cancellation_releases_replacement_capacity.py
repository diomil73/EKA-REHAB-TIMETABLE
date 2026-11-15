from datetime import date, time

from rehab_core.models import (
    DailySessionCancellation,
    Patient,
    Session,
    SessionCancellationKind,
    Therapist,
)
from rehab_core.replacement_options import find_replacement_options
from rehab_core.workload import calculate_therapist_workload


DAY = date(2026, 9, 18)


def _session(session_id: str, patient_id: str, therapist_id: str, at: time) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=at,
        treatment="ΦΘ",
    )


def test_cancelled_occurrence_is_removed_from_provider_workload():
    occupied = _session("S-CANCEL", "P2", "T-FREE", time(9, 0))
    cancellation = DailySessionCancellation(
        cancellation_id="C1",
        target_session_id=occupied.session_id,
        cancellation_date=DAY,
        kind=SessionCancellationKind.PATIENT_NO_SHOW,
        reason="no-show",
    )

    workload = calculate_therapist_workload(
        therapist_id="T-FREE",
        target_date=DAY,
        sessions=[occupied],
        patients=[Patient("P2", "Two")],
        cancellations=[cancellation],
    )

    assert workload.active_sessions == 0
    assert workload.active_timeslots == 0


def test_cancelled_occurrence_releases_exact_slot_for_replacement():
    target = _session("S-TARGET", "P1", "T-ABS", time(9, 0))
    occupied = _session("S-CANCEL", "P2", "T-FREE", time(9, 0))
    cancellation = DailySessionCancellation(
        cancellation_id="C1",
        target_session_id=occupied.session_id,
        cancellation_date=DAY,
        kind=SessionCancellationKind.DEPARTMENT_POSTPONED,
        reason="department postponed",
    )

    options = find_replacement_options(
        target_session=target,
        therapists=[Therapist("T-ABS", "Absent"), Therapist("T-FREE", "Free")],
        sessions=[target, occupied],
        patients=[Patient("P1", "One"), Patient("P2", "Two")],
        cancellations=[cancellation],
        requested_time=time(9, 0),
        timeslots=[time(9, 0)],
    )

    assert len(options) == 1
    option = options[0]
    assert option.provider_id == "T-FREE"
    assert option.exact_time_available is True
    assert option.available_timeslots == (time(9, 0),)
    assert option.active_sessions == 0
    assert option.capacity_remaining == 6
