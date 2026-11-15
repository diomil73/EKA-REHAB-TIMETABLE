from __future__ import annotations

from datetime import time
import json
from pathlib import Path

from rehab_core.service_hours import ServiceOperatingHours, SpecialServiceHoursBook


SCHEMA_VERSION = 1


def _parse_time(value: str) -> time:
    return time.fromisoformat(value)


def load_service_hours(path: str | Path) -> SpecialServiceHoursBook:
    path = Path(path)
    if not path.exists():
        return SpecialServiceHoursBook()

    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != SCHEMA_VERSION:
        raise ValueError("Unsupported service-hours schema")

    services: dict[str, ServiceOperatingHours] = {}
    for item in payload.get("services", []):
        weekdays = {
            int(weekday): tuple(_parse_time(value) for value in values)
            for weekday, values in item.get("allowed_times_by_weekday", {}).items()
        }
        service = ServiceOperatingHours(
            service_id=item["service_id"],
            allowed_times_by_weekday=weekdays,
        )
        services[service.service_id] = service
    return SpecialServiceHoursBook(services=services)


def save_service_hours(book: SpecialServiceHoursBook, path: str | Path) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "schema_version": SCHEMA_VERSION,
        "services": [
            {
                "service_id": service.service_id,
                "allowed_times_by_weekday": {
                    str(weekday): [value.strftime("%H:%M") for value in sorted(values)]
                    for weekday, values in sorted(service.allowed_times_by_weekday.items())
                },
            }
            for service in sorted(book.services.values(), key=lambda value: value.service_id)
        ],
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
