from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable, Optional

from .day_patterns import RehabWeekday, parse_day_pattern
from .models import BaseScheduleEntry


@dataclass(frozen=True)
class PermanentDayProjection:
    weekday: RehabWeekday
    used_timeslots_before: int
    used_timeslots_after: int
    max_timeslots: int
    destination_time: time
    conflicting_entry_ids: tuple[str, ...] = ()

    @property
    def over_capacity(self) -> bool:
        return self.used_timeslots_after > self.max_timeslots

    @property
    def has_conflict(self) -> bool:
        return bool(self.conflicting_entry_ids)

    @property
    def allowed(self) -> bool:
        return not self.over_capacity and not self.has_conflict


@dataclass(frozen=True)
class PermanentAssignmentCheck:
    provider_id: str
    proposed_day_pattern: str
    proposed_time: time
    max_timeslots: int
    source_entry_id: Optional[str]
    days: tuple[PermanentDayProjection, ...]

    @property
    def allowed(self) -> bool:
        return all(day.allowed for day in self.days)

    @property
    def capacity_issues(self) -> tuple[PermanentDayProjection, ...]:
        return tuple(day for day in self.days if day.over_capacity)

    @property
    def conflicts(self) -> tuple[PermanentDayProjection, ...]:
        return tuple(day for day in self.days if day.has_conflict)


def check_permanent_assignment(
    *,
    provider_id: str,
    max_daily_timeslots: int,
    existing_entries: Iterable[BaseScheduleEntry],
    proposed_day_pattern: str,
    proposed_time: time,
    source_entry_id: str | None = None,
    patient_id: str | None = None,
) -> PermanentAssignmentCheck:
    """Validate a permanent recurring assignment across every affected weekday.

    ``source_entry_id`` is excluded before projection. This makes the function
    suitable for permanent therapist/time changes as well as new assignments.

    Capacity is DISTINCT provider clock-times per weekday. A second patient at
    the same provider/time/day is a scheduling conflict even though it would not
    increase the distinct-timeslot count. Complementary day-patterns may share
    the same visual cell because they do not conflict on the same weekday.
    """

    if max_daily_timeslots < 1:
        raise ValueError("max_daily_timeslots must be >= 1")
    proposed_days = parse_day_pattern(proposed_day_pattern)

    entries = tuple(
        entry for entry in existing_entries if entry.base_entry_id != source_entry_id
    )

    occupied: dict[RehabWeekday, set[time]] = {
        weekday: set() for weekday in RehabWeekday
    }
    at_time: dict[RehabWeekday, list[BaseScheduleEntry]] = {
        weekday: [] for weekday in RehabWeekday
    }

    for entry in entries:
        if entry.therapist_id != provider_id:
            continue
        for weekday in parse_day_pattern(entry.day_pattern):
            occupied[weekday].add(entry.start_time)
            if entry.start_time == proposed_time:
                at_time[weekday].append(entry)

    projections: list[PermanentDayProjection] = []
    for weekday in sorted(proposed_days):
        before = len(occupied[weekday])
        after = len(occupied[weekday] | {proposed_time})

        # A different recurring assignment active on the same weekday at the
        # same provider/time is a real collision. An assignment for the same
        # patient is not treated as a collision here; it may represent a split
        # schedule that the higher-level grouping layer can consolidate.
        conflicts = tuple(
            entry.base_entry_id
            for entry in at_time[weekday]
            if patient_id is None or entry.patient_id != patient_id
        )

        projections.append(
            PermanentDayProjection(
                weekday=weekday,
                used_timeslots_before=before,
                used_timeslots_after=after,
                max_timeslots=max_daily_timeslots,
                destination_time=proposed_time,
                conflicting_entry_ids=conflicts,
            )
        )

    return PermanentAssignmentCheck(
        provider_id=provider_id,
        proposed_day_pattern=proposed_day_pattern,
        proposed_time=proposed_time,
        max_timeslots=max_daily_timeslots,
        source_entry_id=source_entry_id,
        days=tuple(projections),
    )


def check_new_assignment(
    *,
    provider_id: str,
    max_daily_timeslots: int,
    existing_entries: Iterable[BaseScheduleEntry],
    day_pattern: str,
    start_time: time,
    patient_id: str | None = None,
) -> PermanentAssignmentCheck:
    """Shared capacity/conflict gate for a new patient assignment."""

    return check_permanent_assignment(
        provider_id=provider_id,
        max_daily_timeslots=max_daily_timeslots,
        existing_entries=existing_entries,
        proposed_day_pattern=day_pattern,
        proposed_time=start_time,
        patient_id=patient_id,
    )
