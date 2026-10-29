from datetime import date, time

from rehab_config.service_hours_store import load_service_hours, save_service_hours
from rehab_core.service_hours import (
    POOL_SERVICE_ID,
    ROBOTIC_SERVICE_ID,
    ServiceOperatingHours,
    SpecialServiceHoursBook,
    filter_times_for_service,
)


def test_service_hours_filter_by_weekday_and_time():
    monday = date(2026, 10, 19)
    book = SpecialServiceHoursBook(
        services={
            ROBOTIC_SERVICE_ID: ServiceOperatingHours(
                ROBOTIC_SERVICE_ID,
                {0: (time(9, 0), time(10, 0))},
            )
        }
    )

    assert filter_times_for_service(
        (time(8, 0), time(9, 0), time(10, 0), time(11, 0)),
        service_id=ROBOTIC_SERVICE_ID,
        target_date=monday,
        hours_book=book,
    ) == (time(9, 0), time(10, 0))


def test_missing_weekday_means_service_closed():
    tuesday = date(2026, 10, 20)
    service = ServiceOperatingHours(POOL_SERVICE_ID, {0: (time(9, 0),)})

    assert service.allowed_times(tuesday) == ()
    assert not service.is_open(tuesday, time(9, 0))


def test_service_hours_round_trip(tmp_path):
    path = tmp_path / "service_hours.json"
    book = SpecialServiceHoursBook(
        services={
            POOL_SERVICE_ID: ServiceOperatingHours(
                POOL_SERVICE_ID,
                {
                    0: (time(9, 30), time(10, 30)),
                    2: (time(11, 0),),
                },
            ),
            ROBOTIC_SERVICE_ID: ServiceOperatingHours(
                ROBOTIC_SERVICE_ID,
                {1: (time(8, 30), time(9, 15))},
            ),
        }
    )

    save_service_hours(book, path)
    loaded = load_service_hours(path)

    assert loaded == book
