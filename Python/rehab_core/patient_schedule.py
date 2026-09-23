from __future__ import annotations

from dataclasses import replace
from datetime import date, time
from typing import Iterable

from .base_schedule import pattern_applies_on_date
from .day_patterns import RehabWeekday, parse_day_pattern
from .models import BaseScheduleEntry, Session
from .replacement_options import ReplacementProviderOption, order_timeslots_by_preference


def recurring_patient_conflict_ids(
    *,
    entries: Iterable[BaseScheduleEntry],
    patient_id: str,
    proposed_day_pattern: str,
    proposed_time: time,
    source_entry_id: str | None = None,
) -> dict[RehabWeekday, tuple[str, ...]]:
    """Return recurring patient conflicts across every proposed weekday.

    This is intentionally specialty-agnostic. Any other treatment for the same
    patient at the same weekday/time is a collision, regardless of whether the
    other treatment has a named provider in PATIENT_PLANNER.
    """

    proposed_days = parse_day_pattern(proposed_day_pattern)
    conflicts: dict[RehabWeekday, list[str]] = {
        weekday: [] for weekday in proposed_days
    }
    for entry in entries:
        if entry.base_entry_id == source_entry_id:
            continue
        if entry.patient_id != patient_id or entry.start_time != proposed_time:
            continue
        active_days = parse_day_pattern(entry.day_pattern)
        for weekday in proposed_days & active_days:
            conflicts[weekday].append(entry.base_entry_id)
    return {
        weekday: tuple(sorted(ids))
        for weekday, ids in conflicts.items()
    }


def nonprovider_patient_conflicts_on_date(
    *,
    entries: Iterable[BaseScheduleEntry],
    patient_id: str,
    target_date: date,
    target_time: time,
) -> tuple[BaseScheduleEntry, ...]:
    """Return non-provider specialty commitments blocking one patient slot.

    Provider-owned sessions (for example ΦΘ/Ρομποτικό) are already represented
    by the operational Session list and are handled by the existing availability
    engine, including replacements and patient absences. This gate adds the
    specialties that previously vanished from availability because they have no
    therapist column in PATIENT_PLANNER, such as Εργο, Λογο, ΕΦΑ, Πισίνα and
    Ανακλινόμενο.
    """

    return tuple(
        entry
        for entry in entries
        if entry.therapist_id is None
        and entry.patient_id == patient_id
        and entry.start_time == target_time
        and pattern_applies_on_date(entry.day_pattern, target_date)
    )


def filter_replacement_options_for_patient_schedule(
    options: Iterable[ReplacementProviderOption],
    *,
    target_session: Session,
    base_entries: Iterable[BaseScheduleEntry],
) -> list[ReplacementProviderOption]:
    """Remove times blocked by another specialty for the same patient."""

    entries = tuple(base_entries)
    if not entries:
        return list(options)

    filtered: list[ReplacementProviderOption] = []
    for option in options:
        allowed = tuple(
            slot
            for slot in option.available_timeslots
            if not nonprovider_patient_conflicts_on_date(
                entries=entries,
                patient_id=target_session.patient_id,
                target_date=target_session.session_date,
                target_time=slot,
            )
        )
        if not allowed:
            continue
        ordered = order_timeslots_by_preference(allowed, option.requested_time)
        filtered.append(
            replace(
                option,
                recommended_time=ordered[0],
                exact_time_available=option.requested_time in ordered,
                available_timeslots=ordered,
            )
        )
    return filtered


def validate_patient_cross_specialty_time(
    *,
    patient_id: str,
    target_date: date,
    target_time: time,
    base_entries: Iterable[BaseScheduleEntry],
) -> None:
    """Hard safety gate before accepting a replacement at one time."""

    conflicts = nonprovider_patient_conflicts_on_date(
        entries=base_entries,
        patient_id=patient_id,
        target_date=target_date,
        target_time=target_time,
    )
    if conflicts:
        treatments = ", ".join(sorted({entry.treatment for entry in conflicts}))
        raise ValueError(
            "Patient has another specialty at selected time: " + treatments
        )
