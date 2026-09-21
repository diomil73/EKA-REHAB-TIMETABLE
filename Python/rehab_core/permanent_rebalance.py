from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable, Mapping

from .day_patterns import RehabWeekday, parse_day_pattern
from .models import BaseScheduleEntry
from .permanent_assignment import PermanentAssignmentCheck, check_permanent_assignment


@dataclass(frozen=True)
class ProviderWeekdayCapacity:
    weekday: RehabWeekday
    used_timeslots: int
    max_timeslots: int

    @property
    def excess(self) -> int:
        return max(0, self.used_timeslots - self.max_timeslots)


@dataclass(frozen=True)
class PermanentRebalanceOption:
    source_entry: BaseScheduleEntry
    destination_provider_id: str
    destination_time: time
    destination_check: PermanentAssignmentCheck
    source_days_freed: tuple[RehabWeekday, ...]
    overcapacity_days_helped: tuple[RehabWeekday, ...]
    source_excess_before: int
    source_excess_after: int
    source_overcapacity_days_before: int
    source_overcapacity_days_after: int

    @property
    def same_time(self) -> bool:
        return self.destination_time == self.source_entry.start_time

    @property
    def fully_resolves_source(self) -> bool:
        return self.source_excess_after == 0

    @property
    def destination_peak_after(self) -> int:
        return max(
            (projection.used_timeslots_after for projection in self.destination_check.days),
            default=0,
        )

    @property
    def destination_total_after(self) -> int:
        return sum(projection.used_timeslots_after for projection in self.destination_check.days)


def _occupied_times_by_weekday(
    entries: Iterable[BaseScheduleEntry],
    *,
    provider_id: str,
    exclude_entry_id: str | None = None,
) -> dict[RehabWeekday, set[time]]:
    occupied = {weekday: set() for weekday in RehabWeekday}
    for entry in entries:
        if entry.base_entry_id == exclude_entry_id:
            continue
        if entry.therapist_id != provider_id:
            continue
        for weekday in parse_day_pattern(entry.day_pattern):
            occupied[weekday].add(entry.start_time)
    return occupied


def provider_weekly_capacity(
    *,
    provider_id: str,
    existing_entries: Iterable[BaseScheduleEntry],
    max_daily_timeslots: int,
    exclude_entry_id: str | None = None,
) -> tuple[ProviderWeekdayCapacity, ...]:
    occupied = _occupied_times_by_weekday(
        existing_entries,
        provider_id=provider_id,
        exclude_entry_id=exclude_entry_id,
    )
    return tuple(
        ProviderWeekdayCapacity(
            weekday=weekday,
            used_timeslots=len(occupied[weekday]),
            max_timeslots=max_daily_timeslots,
        )
        for weekday in RehabWeekday
    )


def _source_days_freed(
    *,
    source_entry: BaseScheduleEntry,
    existing_entries: tuple[BaseScheduleEntry, ...],
) -> tuple[RehabWeekday, ...]:
    if source_entry.therapist_id is None:
        return ()

    remaining = tuple(
        entry
        for entry in existing_entries
        if entry.base_entry_id != source_entry.base_entry_id
        and entry.therapist_id == source_entry.therapist_id
        and entry.start_time == source_entry.start_time
    )
    freed: list[RehabWeekday] = []
    for weekday in sorted(parse_day_pattern(source_entry.day_pattern)):
        still_occupied = any(
            weekday in parse_day_pattern(other.day_pattern) for other in remaining
        )
        if not still_occupied:
            freed.append(weekday)
    return tuple(freed)


def _sum_excess(capacities: Iterable[ProviderWeekdayCapacity]) -> int:
    return sum(item.excess for item in capacities)


def _count_over_days(capacities: Iterable[ProviderWeekdayCapacity]) -> int:
    return sum(1 for item in capacities if item.excess > 0)


def find_permanent_rebalance_options(
    *,
    source_provider_id: str,
    existing_entries: Iterable[BaseScheduleEntry],
    destination_provider_ids: Iterable[str],
    standard_timeslots: Iterable[time],
    provider_capacity_limits: Mapping[str, int] | None = None,
    source_capacity_limit: int = 6,
) -> tuple[PermanentRebalanceOption, ...]:
    """Find capacity-safe permanent moves that actually relieve source overload.

    This is deliberately a planner, not an editor. It never changes the base
    programme. Every option must pass the weekly destination capacity/conflict
    gate and must reduce at least one over-capacity source weekday.
    """

    entries = tuple(existing_entries)
    limits = dict(provider_capacity_limits or {})
    source_before = provider_weekly_capacity(
        provider_id=source_provider_id,
        existing_entries=entries,
        max_daily_timeslots=source_capacity_limit,
    )
    over_before = {item.weekday for item in source_before if item.excess > 0}
    if not over_before:
        return ()

    source_excess_before = _sum_excess(source_before)
    source_over_days_before = _count_over_days(source_before)

    options: list[PermanentRebalanceOption] = []
    source_entries = tuple(
        entry for entry in entries if entry.therapist_id == source_provider_id
    )
    destinations = tuple(
        provider_id
        for provider_id in destination_provider_ids
        if provider_id and provider_id != source_provider_id
    )
    timeslots = tuple(dict.fromkeys(standard_timeslots))

    for source_entry in source_entries:
        freed_days = _source_days_freed(
            source_entry=source_entry,
            existing_entries=entries,
        )
        helped_days = tuple(day for day in freed_days if day in over_before)
        if not helped_days:
            continue

        source_after = provider_weekly_capacity(
            provider_id=source_provider_id,
            existing_entries=entries,
            max_daily_timeslots=source_capacity_limit,
            exclude_entry_id=source_entry.base_entry_id,
        )
        source_excess_after = _sum_excess(source_after)
        source_over_days_after = _count_over_days(source_after)
        if source_excess_after >= source_excess_before:
            continue

        for destination_provider_id in destinations:
            limit = limits.get(destination_provider_id, 6)
            for destination_time in timeslots:
                check = check_permanent_assignment(
                    provider_id=destination_provider_id,
                    max_daily_timeslots=limit,
                    existing_entries=entries,
                    proposed_day_pattern=source_entry.day_pattern,
                    proposed_time=destination_time,
                    source_entry_id=source_entry.base_entry_id,
                    patient_id=source_entry.patient_id,
                )
                if not check.allowed:
                    continue
                options.append(
                    PermanentRebalanceOption(
                        source_entry=source_entry,
                        destination_provider_id=destination_provider_id,
                        destination_time=destination_time,
                        destination_check=check,
                        source_days_freed=freed_days,
                        overcapacity_days_helped=helped_days,
                        source_excess_before=source_excess_before,
                        source_excess_after=source_excess_after,
                        source_overcapacity_days_before=source_over_days_before,
                        source_overcapacity_days_after=source_over_days_after,
                    )
                )

    def sort_key(option: PermanentRebalanceOption):
        # Prefer moves that completely fix the source, then same clock time,
        # then lower destination peak/weekly footprint. Deterministic text/time
        # tiebreakers keep CLI output stable for regression tests.
        return (
            0 if option.fully_resolves_source else 1,
            -len(option.overcapacity_days_helped),
            0 if option.same_time else 1,
            option.destination_peak_after,
            option.destination_total_after,
            option.destination_provider_id.casefold(),
            option.destination_time,
            option.source_entry.start_time,
            option.source_entry.patient_id,
        )

    return tuple(sorted(options, key=sort_key))
