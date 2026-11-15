from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, time
from typing import Iterable, Mapping


ROBOTIC_SERVICE_ID = "robotic"
POOL_SERVICE_ID = "pool"


@dataclass(frozen=True)
class ServiceOperatingHours:
    """Persistent allowed clock slots for one special service.

    Weekdays use Python's Monday=0 ... Sunday=6 convention. An omitted weekday
    means the service is closed that day. The contract is intentionally about
    the service/resource, not about any individual provider.
    """

    service_id: str
    allowed_times_by_weekday: Mapping[int, tuple[time, ...]] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.service_id.strip():
            raise ValueError("service_id is required")
        for weekday, values in self.allowed_times_by_weekday.items():
            if weekday < 0 or weekday > 6:
                raise ValueError("weekday must be between 0 and 6")
            if len(set(values)) != len(values):
                raise ValueError("allowed times must be unique per weekday")

    def allowed_times(self, target_date: date) -> tuple[time, ...]:
        return tuple(sorted(self.allowed_times_by_weekday.get(target_date.weekday(), ())))

    def is_open(self, target_date: date, target_time: time) -> bool:
        return target_time in self.allowed_times(target_date)


@dataclass(frozen=True)
class SpecialServiceHoursBook:
    services: Mapping[str, ServiceOperatingHours] = field(default_factory=dict)

    def allowed_times(self, service_id: str, target_date: date) -> tuple[time, ...]:
        service = self.services.get(service_id)
        if service is None:
            return ()
        return service.allowed_times(target_date)

    def is_open(self, service_id: str, target_date: date, target_time: time) -> bool:
        service = self.services.get(service_id)
        return bool(service and service.is_open(target_date, target_time))


def filter_times_for_service(
    times: Iterable[time],
    *,
    service_id: str,
    target_date: date,
    hours_book: SpecialServiceHoursBook,
) -> tuple[time, ...]:
    allowed = set(hours_book.allowed_times(service_id, target_date))
    return tuple(sorted({value for value in times if value in allowed}))
