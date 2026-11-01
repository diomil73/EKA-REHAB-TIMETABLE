from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import datetime
from typing import Iterable

from .daily_state import DailySessionState
from .day_finalization import (
    FinalizedDaySnapshot,
    FinalizedSnapshotStatus,
    build_finalized_day_snapshot,
)
from .day_history_events import DayHistoryEvent, DayHistoryEventType


class DayAmendmentError(ValueError):
    """Raised when a closed-day amendment/version transition is invalid."""


@dataclass(frozen=True)
class DayAmendmentPlan:
    previous_snapshot: FinalizedDaySnapshot
    superseded_snapshot: FinalizedDaySnapshot
    replacement_snapshot: FinalizedDaySnapshot
    started_event: DayHistoryEvent
    finalized_event: DayHistoryEvent


def _require_timezone(value: datetime, *, field: str) -> None:
    if value.tzinfo is None or value.utcoffset() is None:
        raise DayAmendmentError(f"{field} must include a timezone offset")


def build_day_amendment_plan(
    states: Iterable[DailySessionState],
    *,
    previous_snapshot: FinalizedDaySnapshot,
    replacement_snapshot_id: str,
    started_event_id: str,
    finalized_event_id: str,
    actor_or_source: str,
    amendment_started_at: datetime,
    amendment_finalized_at: datetime,
    reason_code: str,
    summary: str,
    planned_session_count: int | None = None,
    approved_exception_count: int = 0,
) -> DayAmendmentPlan:
    """Build a validated append-only amendment transition for one closed day.

    The previous snapshot is never mutated in-place. The returned plan contains
    an explicit SUPERSEDED representation of the old version, a new
    CURRENT_FINAL version that supersedes it, and linked STARTED/FINALIZED audit
    events. Workbook persistence is handled by a separate integration layer.
    """

    if previous_snapshot.status != FinalizedSnapshotStatus.CURRENT_FINAL:
        raise DayAmendmentError("Previous snapshot must be CURRENT_FINAL")
    if previous_snapshot.unresolved_sessions != 0:
        raise DayAmendmentError("Previous finalized snapshot cannot contain unresolved sessions")

    replacement_snapshot_id = replacement_snapshot_id.strip()
    started_event_id = started_event_id.strip()
    finalized_event_id = finalized_event_id.strip()
    actor = actor_or_source.strip()
    reason = reason_code.strip()
    summary_text = summary.strip()

    if not replacement_snapshot_id:
        raise DayAmendmentError("replacement_snapshot_id is required")
    if replacement_snapshot_id == previous_snapshot.snapshot_id:
        raise DayAmendmentError("Replacement snapshot must have a new SnapshotId")
    if not started_event_id or not finalized_event_id:
        raise DayAmendmentError("Both amendment event IDs are required")
    if started_event_id == finalized_event_id:
        raise DayAmendmentError("Amendment event IDs must be different")
    if not actor:
        raise DayAmendmentError("actor_or_source is required")
    if not reason:
        raise DayAmendmentError("reason_code is required")
    if not summary_text:
        raise DayAmendmentError("summary is required")

    _require_timezone(amendment_started_at, field="amendment_started_at")
    _require_timezone(amendment_finalized_at, field="amendment_finalized_at")
    if amendment_finalized_at < amendment_started_at:
        raise DayAmendmentError("amendment_finalized_at cannot be earlier than amendment_started_at")

    replacement = build_finalized_day_snapshot(
        states,
        snapshot_id=replacement_snapshot_id,
        service_date=previous_snapshot.service_date,
        finalized_at=amendment_finalized_at,
        version=previous_snapshot.version + 1,
        supersedes_snapshot_id=previous_snapshot.snapshot_id,
        planned_session_count=(
            previous_snapshot.planned_session_count
            if planned_session_count is None
            else planned_session_count
        ),
        approved_exception_count=approved_exception_count,
    )

    superseded = replace(
        previous_snapshot,
        status=FinalizedSnapshotStatus.SUPERSEDED,
    )

    started = DayHistoryEvent(
        event_id=started_event_id,
        service_date=previous_snapshot.service_date,
        event_type=DayHistoryEventType.DAY_AMENDMENT_STARTED,
        recorded_at=amendment_started_at,
        actor_or_source=actor,
        reason_code=reason,
        related_snapshot_id=previous_snapshot.snapshot_id,
        summary=summary_text,
    )
    finalized = DayHistoryEvent(
        event_id=finalized_event_id,
        service_date=previous_snapshot.service_date,
        event_type=DayHistoryEventType.DAY_AMENDMENT_FINALIZED,
        recorded_at=amendment_finalized_at,
        actor_or_source=actor,
        reason_code=reason,
        related_event_id=started_event_id,
        related_snapshot_id=replacement.snapshot_id,
        summary=(
            f"{summary_text} | superseded {previous_snapshot.snapshot_id} "
            f"with {replacement.snapshot_id}"
        ),
    )

    return DayAmendmentPlan(
        previous_snapshot=previous_snapshot,
        superseded_snapshot=superseded,
        replacement_snapshot=replacement,
        started_event=started,
        finalized_event=finalized,
    )
