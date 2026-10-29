from datetime import date, datetime, timezone

import pytest

from rehab_core.day_history_events import (
    DayHistoryEventError,
    DayHistoryEventType,
    build_day_finalized_event,
)


DAY = date(2026, 10, 26)
STAMP = datetime(2026, 10, 26, 16, 0, tzinfo=timezone.utc)


def test_build_day_finalized_event_links_snapshot_and_is_append_only_fact():
    event = build_day_finalized_event(
        event_id="finalize:2026-10-26:v1",
        service_date=DAY,
        recorded_at=STAMP,
        actor_or_source="close-day-action",
        snapshot_id="finalized:2026-10-26:v1",
    )

    assert event.event_type == DayHistoryEventType.DAY_FINALIZED
    assert event.related_snapshot_id == "finalized:2026-10-26:v1"
    assert event.recorded_at == STAMP
    assert event.actor_or_source == "close-day-action"


def test_build_day_finalized_event_rejects_naive_timestamp():
    with pytest.raises(DayHistoryEventError, match="timezone"):
        build_day_finalized_event(
            event_id="event-1",
            service_date=DAY,
            recorded_at=datetime(2026, 10, 26, 16, 0),
            actor_or_source="close-day-action",
            snapshot_id="snapshot-1",
        )


def test_build_day_finalized_event_requires_identity_fields():
    with pytest.raises(DayHistoryEventError, match="event_id"):
        build_day_finalized_event(
            event_id=" ",
            service_date=DAY,
            recorded_at=STAMP,
            actor_or_source="close-day-action",
            snapshot_id="snapshot-1",
        )

    with pytest.raises(DayHistoryEventError, match="actor_or_source"):
        build_day_finalized_event(
            event_id="event-1",
            service_date=DAY,
            recorded_at=STAMP,
            actor_or_source=" ",
            snapshot_id="snapshot-1",
        )
