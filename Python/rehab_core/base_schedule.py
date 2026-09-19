from __future__ import annotations

from datetime import date
from typing import Iterable

from .day_patterns import RehabWeekday, parse_day_pattern
from .models import BaseScheduleEntry, Session


def pattern_applies_on_date(pattern: str, target_date: date) -> bool:
    """Return whether a recurring workbook pattern applies on target_date."""

    try:
        weekday = RehabWeekday(target_date.weekday())
    except ValueError:
        return False
    return weekday in parse_day_pattern(pattern)


def materialize_sessions_for_date(
    entries: Iterable[BaseScheduleEntry], target_date: date
) -> list[Session]:
    """Turn recurring base entries into dated provider sessions.

    Entries without a therapist are intentionally skipped because the current
    replacement engine only operates on provider-owned sessions. They remain
    available to the Excel adapter/audit and are not discarded there.
    """

    sessions: list[Session] = []
    for entry in entries:
        if entry.therapist_id is None:
            continue
        if not pattern_applies_on_date(entry.day_pattern, target_date):
            continue
        sessions.append(
            Session(
                session_id=f"{entry.base_entry_id}@{target_date.isoformat()}",
                patient_id=entry.patient_id,
                therapist_id=entry.therapist_id,
                session_date=target_date,
                start_time=entry.start_time,
                treatment=entry.treatment,
                robotic=entry.robotic,
            )
        )
    return sessions
