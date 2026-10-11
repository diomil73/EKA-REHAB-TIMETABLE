from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta, time
from pathlib import Path
from typing import Iterable

from rehab_core.day_patterns import parse_day_pattern

from .reader import read_settings
from .workbook_scheduling_candidates import workbook_therapist_candidates


@dataclass(frozen=True)
class RecurringTherapistChoice:
    therapist_id: str
    display_name: str
    workload_count: int
    available_slots: tuple[time, ...]

    @property
    def workload_label(self) -> str:
        return f"{self.display_name} {self.workload_count}"


def _normalize(value: str) -> str:
    return value.strip().casefold()


def _pattern_dates(day_pattern: str, anchor_date: date) -> tuple[date, ...]:
    weekdays = tuple(sorted({int(day) for day in parse_day_pattern(day_pattern)}))
    if not weekdays:
        raise ValueError(f"Day pattern {day_pattern!r} resolves to no weekdays")

    dates: list[date] = []
    for weekday in weekdays:
        offset = (weekday - anchor_date.weekday()) % 7
        dates.append(anchor_date + timedelta(days=offset))
    return tuple(dates)


def recurring_therapist_choices(
    workbook_path: str | Path,
    *,
    patient_id: str,
    treatment: str,
    day_pattern: str,
    anchor_date: date,
    additional_robotic_capable_names: Iterable[str] = (),
) -> tuple[RecurringTherapistChoice, ...]:
    """Return therapists/slots feasible on every weekday in a recurring pattern.

    SETTINGS remains authoritative for the accepted day-pattern label. For each
    weekday in the pattern we reuse the existing concrete-day scheduling engine,
    then intersect therapist availability and time slots. A therapist is offered
    only if the same start time is conflict-free for the patient and therapist on
    every representative weekday of the recurrence.
    """

    settings = read_settings(workbook_path)
    patterns = {_normalize(item): item for item in settings.day_patterns}
    key = _normalize(day_pattern)
    if key not in patterns:
        raise ValueError(f"Day pattern {day_pattern!r} is not defined in SETTINGS")
    canonical_pattern = patterns[key]

    per_date = []
    for target_date in _pattern_dates(canonical_pattern, anchor_date):
        per_date.append(
            workbook_therapist_candidates(
                workbook_path,
                patient_id=patient_id,
                treatment=treatment,
                target_date=target_date,
                additional_robotic_capable_names=additional_robotic_capable_names,
            )
        )

    if not per_date:
        return ()

    by_date = [
        {candidate.therapist_id: candidate for candidate in candidates}
        for candidates in per_date
    ]
    common_ids = set(by_date[0])
    for mapping in by_date[1:]:
        common_ids.intersection_update(mapping)

    result: list[RecurringTherapistChoice] = []
    for therapist_id in common_ids:
        candidates = [mapping[therapist_id] for mapping in by_date]
        common_slots = set(candidates[0].available_slots)
        for candidate in candidates[1:]:
            common_slots.intersection_update(candidate.available_slots)
        if not common_slots:
            continue
        result.append(
            RecurringTherapistChoice(
                therapist_id=therapist_id,
                display_name=candidates[0].display_name,
                workload_count=sum(candidate.active_sessions for candidate in candidates),
                available_slots=tuple(sorted(common_slots)),
            )
        )

    result.sort(
        key=lambda item: (
            item.workload_count,
            item.display_name.casefold(),
            item.therapist_id,
        )
    )
    return tuple(result)
