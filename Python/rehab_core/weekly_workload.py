from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Iterable

from .workload import TherapistWorkload


@dataclass(frozen=True)
class TherapistWeeklyWorkload:
    therapist_id: str
    week_start: date
    daily_timeslots: tuple[int, int, int, int, int]
    total_timeslots: int
    average_timeslots: float
    peak_timeslots: int
    days_at_capacity: int

    @property
    def week_end(self) -> date:
        return self.week_start + timedelta(days=4)


def monday_of_week(value: date) -> date:
    """Return the Monday anchoring the ISO work week containing ``value``."""

    return value - timedelta(days=value.weekday())


def build_therapist_weekly_workload(
    therapist_id: str,
    week_start: date,
    daily_workloads: Iterable[TherapistWorkload],
    *,
    capacity: int = 6,
) -> TherapistWeeklyWorkload:
    """Aggregate persisted daily therapist workloads into a Monday-Friday history row.

    The input is deliberately a sequence of dated operational workload facts rather
    than a live worksheet projection. Missing weekdays are represented as zero.
    Capacity is timeslot-based, matching ``TherapistWorkload.active_timeslots``.
    """

    if capacity < 1:
        raise ValueError("capacity must be >= 1")

    anchor = monday_of_week(week_start)
    by_date: dict[date, int] = {}
    for workload in daily_workloads:
        if workload.therapist_id != therapist_id:
            continue
        if not anchor <= workload.workload_date <= anchor + timedelta(days=4):
            continue
        by_date[workload.workload_date] = workload.active_timeslots

    daily = tuple(
        by_date.get(anchor + timedelta(days=offset), 0)
        for offset in range(5)
    )
    total = sum(daily)
    return TherapistWeeklyWorkload(
        therapist_id=therapist_id,
        week_start=anchor,
        daily_timeslots=daily,  # type: ignore[arg-type]
        total_timeslots=total,
        average_timeslots=total / 5,
        peak_timeslots=max(daily, default=0),
        days_at_capacity=sum(value >= capacity for value in daily),
    )


def build_weekly_workload_history(
    daily_workloads: Iterable[TherapistWorkload],
    *,
    capacity: int = 6,
) -> tuple[TherapistWeeklyWorkload, ...]:
    """Build deterministic therapist/week history from dated workload facts."""

    workloads = tuple(daily_workloads)
    keys = sorted(
        {
            (workload.therapist_id, monday_of_week(workload.workload_date))
            for workload in workloads
        },
        key=lambda item: (item[1], item[0].casefold()),
    )
    return tuple(
        build_therapist_weekly_workload(
            therapist_id,
            week_start,
            workloads,
            capacity=capacity,
        )
        for therapist_id, week_start in keys
    )
