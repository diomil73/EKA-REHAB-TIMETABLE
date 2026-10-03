from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from enum import IntEnum
from typing import Iterable, Sequence

from rehab_core.day_patterns import parse_day_pattern, patterns_overlap
from rehab_core.models import BaseScheduleEntry


class SuggestionMode(IntEnum):
    CHANGE_TIME = 1
    CHANGE_DAYS = 2
    CHANGE_TIME_AND_DAYS = 3
    CHANGE_THERAPIST = 4
    CHANGE_THERAPIST_AND_TIME = 5
    CHANGE_THERAPIST_AND_DAYS = 6
    CHANGE_THERAPIST_TIME_AND_DAYS = 7


SUGGESTION_MODE_LABELS: tuple[str, ...] = (
    "Αλλαγή ώρας",
    "Αλλαγή ημερών",
    "Αλλαγή ώρας και ημερών",
    "Αλλαγή θεραπευτή",
    "Αλλαγή θεραπευτή και ώρας",
    "Αλλαγή θεραπευτή και ημερών",
    "Αλλαγή θεραπευτή, ώρας και ημερών",
)


@dataclass(frozen=True)
class OutpatientScheduleSuggestion:
    therapist_id: str
    start_time: time
    day_pattern: str


def _minutes(value: time) -> int:
    return value.hour * 60 + value.minute


def _provider_slot_is_free(
    *,
    therapist_id: str,
    start_time: time,
    day_pattern: str,
    existing_entries: Iterable[BaseScheduleEntry],
) -> bool:
    provider = therapist_id.strip().casefold()
    if not provider:
        return False

    for entry in existing_entries:
        if not entry.therapist_id:
            continue
        if entry.therapist_id.strip().casefold() != provider:
            continue
        if entry.start_time != start_time:
            continue
        if patterns_overlap(day_pattern, entry.day_pattern):
            return False
    return True


def _changed_dimensions(
    *,
    original_therapist: str,
    original_time: time,
    original_days: str,
    therapist_id: str,
    start_time: time,
    day_pattern: str,
) -> tuple[bool, bool, bool]:
    therapist_changed = therapist_id.strip().casefold() != original_therapist.strip().casefold()
    time_changed = start_time != original_time
    days_changed = parse_day_pattern(day_pattern) != parse_day_pattern(original_days)
    return therapist_changed, time_changed, days_changed


def _mode_matches(
    mode: SuggestionMode,
    *,
    therapist_changed: bool,
    time_changed: bool,
    days_changed: bool,
) -> bool:
    expected = {
        SuggestionMode.CHANGE_TIME: (False, True, False),
        SuggestionMode.CHANGE_DAYS: (False, False, True),
        SuggestionMode.CHANGE_TIME_AND_DAYS: (False, True, True),
        SuggestionMode.CHANGE_THERAPIST: (True, False, False),
        SuggestionMode.CHANGE_THERAPIST_AND_TIME: (True, True, False),
        SuggestionMode.CHANGE_THERAPIST_AND_DAYS: (True, False, True),
        SuggestionMode.CHANGE_THERAPIST_TIME_AND_DAYS: (True, True, True),
    }
    return expected[mode] == (therapist_changed, time_changed, days_changed)


def suggest_outpatient_schedule_slots(
    *,
    mode: SuggestionMode | int,
    preferred_therapist: str,
    preferred_time: time,
    preferred_day_pattern: str,
    therapist_names: Sequence[str],
    timeslots: Sequence[time],
    day_patterns: Sequence[str],
    existing_entries: Iterable[BaseScheduleEntry],
    limit: int = 5,
) -> tuple[OutpatientScheduleSuggestion, ...]:
    """Return safe schedule alternatives matching exactly the selected change mode.

    A therapist/time/day candidate is considered safe when that therapist has no
    recurring assignment at the same time on any overlapping weekday. Suggestions
    are ranked by least disruption: closest time first, then smallest weekday
    difference, while preserving workbook order for therapists/day patterns as a
    stable tie-breaker.
    """

    selected_mode = SuggestionMode(mode)
    if limit < 1:
        return ()

    preferred_provider = preferred_therapist.strip()
    if not preferred_provider:
        raise ValueError("preferred_therapist is required for schedule suggestions")

    preferred_days = parse_day_pattern(preferred_day_pattern)
    entries = tuple(existing_entries)

    candidates: list[tuple[tuple[int, int, int, int, str], OutpatientScheduleSuggestion]] = []
    seen: set[tuple[str, time, frozenset]] = set()

    for therapist_index, therapist in enumerate(therapist_names):
        therapist_text = therapist.strip()
        if not therapist_text:
            continue
        for candidate_time in timeslots:
            for day_index, candidate_days in enumerate(day_patterns):
                parsed_days = parse_day_pattern(candidate_days)
                key = (therapist_text.casefold(), candidate_time, parsed_days)
                if key in seen:
                    continue
                seen.add(key)

                changed = _changed_dimensions(
                    original_therapist=preferred_provider,
                    original_time=preferred_time,
                    original_days=preferred_day_pattern,
                    therapist_id=therapist_text,
                    start_time=candidate_time,
                    day_pattern=candidate_days,
                )
                if not _mode_matches(
                    selected_mode,
                    therapist_changed=changed[0],
                    time_changed=changed[1],
                    days_changed=changed[2],
                ):
                    continue

                if not _provider_slot_is_free(
                    therapist_id=therapist_text,
                    start_time=candidate_time,
                    day_pattern=candidate_days,
                    existing_entries=entries,
                ):
                    continue

                time_distance = abs(_minutes(candidate_time) - _minutes(preferred_time))
                day_distance = len(preferred_days.symmetric_difference(parsed_days))
                score = (
                    time_distance,
                    day_distance,
                    therapist_index,
                    day_index,
                    therapist_text.casefold(),
                )
                candidates.append(
                    (
                        score,
                        OutpatientScheduleSuggestion(
                            therapist_id=therapist_text,
                            start_time=candidate_time,
                            day_pattern=candidate_days,
                        ),
                    )
                )

    candidates.sort(key=lambda item: item[0])
    return tuple(item[1] for item in candidates[:limit])
