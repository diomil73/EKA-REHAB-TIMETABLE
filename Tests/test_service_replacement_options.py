from datetime import date, time

from rehab_core.models import ReplacementProviderKind, Session
from rehab_core.replacement_options import ReplacementProviderOption
from rehab_core.service_hours import (
    POOL_SERVICE_ID,
    ROBOTIC_SERVICE_ID,
    ServiceOperatingHours,
    SpecialServiceHoursBook,
)
from rehab_core.service_replacement_options import (
    filter_replacement_options_for_service_hours,
    service_id_for_session,
)


TODAY = date(2026, 10, 19)


def _option(*times: time) -> ReplacementProviderOption:
    return ReplacementProviderOption(
        provider_id="T1",
        display_name="Therapist 1",
        provider_kind=ReplacementProviderKind.THERAPIST,
        active_sessions=0,
        active_timeslots=0,
        infectious_sessions=0,
        robotic_sessions=0,
        replacement_sessions=0,
        requested_time=times[0],
        recommended_time=times[0],
        exact_time_available=True,
        available_timeslots=times,
    )


def test_robotic_session_uses_robotic_service_hours_automatically():
    session = Session("S1", "P1", "T0", TODAY, time(9, 0), robotic=True)
    book = SpecialServiceHoursBook(
        services={
            ROBOTIC_SERVICE_ID: ServiceOperatingHours(
                ROBOTIC_SERVICE_ID,
                {0: (time(10, 0),)},
            )
        }
    )

    filtered = filter_replacement_options_for_service_hours(
        [_option(time(9, 0), time(10, 0), time(11, 0))],
        target_session=session,
        hours_book=book,
    )

    assert filtered[0].available_timeslots == (time(10, 0),)
    assert filtered[0].recommended_time == time(10, 0)
    assert filtered[0].exact_time_available is False


def test_explicit_pool_service_filters_without_guessing_treatment_text():
    session = Session("S1", "P1", "T0", TODAY, time(9, 0), treatment="Anything")
    book = SpecialServiceHoursBook(
        services={
            POOL_SERVICE_ID: ServiceOperatingHours(
                POOL_SERVICE_ID,
                {0: (time(11, 0),)},
            )
        }
    )

    filtered = filter_replacement_options_for_service_hours(
        [_option(time(9, 0), time(11, 0))],
        target_session=session,
        hours_book=book,
        explicit_service_id=POOL_SERVICE_ID,
    )

    assert filtered[0].available_timeslots == (time(11, 0),)


def test_closed_service_removes_provider_option_entirely():
    session = Session("S1", "P1", "T0", TODAY, time(9, 0), robotic=True)
    book = SpecialServiceHoursBook(
        services={ROBOTIC_SERVICE_ID: ServiceOperatingHours(ROBOTIC_SERVICE_ID, {})}
    )

    assert filter_replacement_options_for_service_hours(
        [_option(time(9, 0), time(10, 0))],
        target_session=session,
        hours_book=book,
    ) == []


def test_normal_session_is_unchanged_without_explicit_service():
    session = Session("S1", "P1", "T0", TODAY, time(9, 0))
    options = [_option(time(9, 0), time(10, 0))]

    assert service_id_for_session(session) is None
    assert filter_replacement_options_for_service_hours(
        options,
        target_session=session,
        hours_book=SpecialServiceHoursBook(),
    ) == options
