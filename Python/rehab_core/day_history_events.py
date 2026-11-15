from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from enum import Enum


class DayHistoryEventType(str, Enum):
    DAY_FINALIZED = "DAY_FINALIZED"
    DAY_AMENDMENT_STARTED = "DAY_AMENDMENT_STARTED"
    DAY_AMENDMENT_FINALIZED = "DAY_AMENDMENT_FINALIZED"


class DayHistoryEventError(ValueError):
    """Raised when an append-only historical event is invalid."""


@dataclass(frozen=True)
class DayHistoryEvent:
    event_id: str
    service_date: date
    event_type: DayHistoryEventType
    recorded_at: datetime
    actor_or_source: str
    therapist_id: str | None = None
    patient_id: str | None = None
    session_id: str | None = None
    reason_code: str | None = None
    related_event_id: str | None = None
    related_snapshot_id: str | None = None
    summary: str = ""


def build_day_finalized_event(
    *,
    event_id: str,
    service_date: date,
    recorded_at: datetime,
    actor_or_source: str,
    snapshot_id: str,
    summary: str | None = None,
) -> DayHistoryEvent:
    event_id = event_id.strip()
    actor_or_source = actor_or_source.strip()
    snapshot_id = snapshot_id.strip()
    if not event_id:
        raise DayHistoryEventError("event_id is required")
    if not actor_or_source:
        raise DayHistoryEventError("actor_or_source is required")
    if not snapshot_id:
        raise DayHistoryEventError("snapshot_id is required")
    if recorded_at.tzinfo is None or recorded_at.utcoffset() is None:
        raise DayHistoryEventError("recorded_at must include a timezone offset")

    return DayHistoryEvent(
        event_id=event_id,
        service_date=service_date,
        event_type=DayHistoryEventType.DAY_FINALIZED,
        recorded_at=recorded_at,
        actor_or_source=actor_or_source,
        related_snapshot_id=snapshot_id,
        summary=(
            summary.strip()
            if summary and summary.strip()
            else f"Day finalized into snapshot {snapshot_id}"
        ),
    )
