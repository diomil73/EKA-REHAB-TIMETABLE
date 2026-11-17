"""Lossless candidate mapping of clear-column weekly treatment cells.

No workbook writes, no patient names, no calendar dates, no automatic admission.
Do not compress differing weekday slots back into one PATIENT_PLANNER time cell.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from decimal import Decimal
from typing import Any
import re
import unicodedata


class ProgrammeMappingError(ValueError):
    """A source segment is unsafe to assign without operator review."""


WEEKDAYS = ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")
ALIASES = {
    "Δ": "ΔΕ", "ΔΕ": "ΔΕ", "ΔΕΥ": "ΔΕ", "ΔΕΥΤ": "ΔΕ", "ΔΕΥΤΕΡΑ": "ΔΕ",
    "Τ": "ΤΡ", "ΤΡ": "ΤΡ", "ΤΡΙ": "ΤΡ", "ΤΡΙΤ": "ΤΡ", "ΤΡΙΤΗ": "ΤΡ",
    "ΤΕ": "ΤΕ", "ΤΕΤ": "ΤΕ", "ΤΕΤΑΡΤΗ": "ΤΕ",
    "ΠΕ": "ΠΕ", "ΠΕΜ": "ΠΕ", "ΠΕΜΠ": "ΠΕ", "ΠΕΜΠΤΗ": "ΠΕ",
    "Π": "ΠΑ", "ΠΑ": "ΠΑ", "ΠΑΡ": "ΠΑ", "ΠΑΡΑ": "ΠΑ", "ΠΑΡΑΣΚΕΥΗ": "ΠΑ",
}


@dataclass(frozen=True)
class WeekdaySlot:
    weekday: str
    start_time: time
    therapist_name: str | None


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _day_text(value: Any) -> str:
    normalized = unicodedata.normalize("NFD", _text(value).upper())
    text = "".join(x for x in normalized if unicodedata.category(x) != "Mn")
    return (
        text.replace("E", "Ε").replace("A", "Α").replace("–", "-")
        .replace("—", "-").replace("(", "").replace(")", "")
        .replace(",", "-").replace(" ", "-")
    )


def parse_clock(value: Any) -> time:
    """Interpret native Excel fractional days and typed HH.MM/HH:MM correctly."""
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, (float, int, Decimal)) and not isinstance(value, bool):
        number = float(value)
        if 0 <= number < 1:
            minutes = round(number * 1440)
            if minutes >= 1440:
                raise ProgrammeMappingError("Excel time wraps beyond the working day")
            return time(minutes // 60, minutes % 60)
        if isinstance(value, int) and 0 <= value <= 23:
            return time(value, 0)
    raw = _text(value).replace(",", ".").replace("..", ".")
    if not raw:
        raise ProgrammeMappingError("Missing clock time")
    if re.fullmatch(r"\d{1,2}", raw):
        hour = int(raw)
        if hour < 24:
            return time(hour)
    # Greek keyboard tonos is accepted only between hour and two minute digits.
    if re.fullmatch(r"[0-9]{1,2}\u0384[0-9]{2}", raw):
        raw = raw.replace("\u0384", ":")
    match = re.fullmatch(r"(\d{1,2})[.:](\d{1,2})", raw)
    if match:
        hour, minute = int(match.group(1)), int(match.group(2))
        # Excel float 11.3 represents the typed time 11.30.
        if "." in raw and len(match.group(2)) == 1:
            minute *= 10
        if hour < 24 and minute < 60:
            return time(hour, minute)
    raise ProgrammeMappingError(f"Unrecognized clock: {raw!r}")


def parse_weekday_groups(value: Any, *, implicit_daily: bool = False,
                         allow_cross_group_repeat: bool = False) -> tuple[tuple[str, ...], ...]:
    raw = _day_text(value)
    if not raw:
        if implicit_daily:
            return (WEEKDAYS,)
        raise ProgrammeMappingError("Missing explicit weekday")
    if raw in {"ΚΑΘΗΜΕΡΙΝΑ", "ΚΑΘΗΜΕΡΙΝΟ", "ΚΑΘ/ΝΑ", "ΚΑΘ-ΝΑ", "ΚΑΘΕ-ΜΕΡΑ"}:
        return (WEEKDAYS,)
    groups = []
    seen = set()
    for group in raw.split("/"):
        days = []
        for token in filter(None, group.split("-")):
            if token not in ALIASES:
                raise ProgrammeMappingError(f"Unknown weekday token: {token!r}")
            weekday = ALIASES[token]
            if weekday in days or (weekday in seen and not allow_cross_group_repeat):
                raise ProgrammeMappingError(f"Duplicate weekday: {weekday}")
            seen.add(weekday)
            days.append(weekday)
        if not days:
            raise ProgrammeMappingError("Empty weekday group")
        groups.append(tuple(days))
    return tuple(groups)


def split_weekly_cell(
    *,
    clock: Any,
    weekdays: Any,
    therapist: Any = None,
    requires_therapist: bool = False,
    implicit_daily: bool = False,
    allow_distinct_same_day: bool = False,
) -> tuple[WeekdaySlot, ...]:
    """Map slash-separated fields positionally; do not guess mismatched cardinalities.

    A single clock/name is broadcast to every day group. Two clocks require
    exactly two slash-separated *day groups*, never an inferred pairing.
    """
    if not _text(clock):
        raise ProgrammeMappingError("Missing clock time")
    raw_clocks = _text(clock).split("/") if isinstance(clock, str) else [clock]
    parsed_times = [parse_clock(value) for value in raw_clocks]
    groups = parse_weekday_groups(weekdays, implicit_daily=implicit_daily,
                                  allow_cross_group_repeat=allow_distinct_same_day)
    names = [x.strip() for x in _text(therapist).split("/")] if _text(therapist) else []
    if requires_therapist and (not names or any(not name for name in names)):
        raise ProgrammeMappingError("Named therapist required for ΦΘ / Ρομποτική")
    group_count = len(groups)
    if len(parsed_times) not in (1, group_count):
        raise ProgrammeMappingError("Time/day groups are not positionally aligned")
    if names and len(names) not in (1, group_count):
        raise ProgrammeMappingError("Provider/day groups are not positionally aligned")
    result = []
    seen_slots = set()
    for index, days in enumerate(groups):
        clock_value = parsed_times[0] if len(parsed_times) == 1 else parsed_times[index]
        name = (names[0] if len(names) == 1 else names[index]) if names else None
        for weekday in days:
            key = (weekday, clock_value)
            if key in seen_slots:
                raise ProgrammeMappingError("Duplicate weekday/time slot")
            seen_slots.add(key)
            result.append(WeekdaySlot(weekday, clock_value, name))
    return tuple(result)


@dataclass(frozen=True)
class SourceWeeklyEntry:
    """One preserved weekday placement originating from a specific source cell."""
    patient_id: str
    treatment: str
    weekday: str
    start_time: time
    therapist_name: str | None
    source_cell: str

    def __post_init__(self) -> None:
        if not self.patient_id.strip() or not self.treatment.strip():
            raise ProgrammeMappingError("PatientID and treatment are required")
        if self.weekday not in WEEKDAYS:
            raise ProgrammeMappingError("Unknown canonical weekday")
        if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", self.source_cell):
            raise ProgrammeMappingError("Exact source cell address is required")


def make_weekly_entries(
    *,
    patient_id: str,
    treatment: str,
    source_cell: str,
    clock: Any,
    weekdays: Any,
    therapist: Any = None,
    implicit_daily: bool = False,
) -> tuple[SourceWeeklyEntry, ...]:
    """Prepare a lossless recurrence map, never an appointment or history fact."""
    if not patient_id.strip() or not treatment.strip():
        raise ProgrammeMappingError("PatientID and treatment are required")
    required = treatment in {"ΦΘ", "Ρομποτικό"}
    slots = split_weekly_cell(
        clock=clock, weekdays=weekdays, therapist=therapist,
        requires_therapist=required, implicit_daily=implicit_daily,
        allow_distinct_same_day=treatment == "ΕΦΑ",
    )
    return tuple(
        SourceWeeklyEntry(
            patient_id, treatment, item.weekday, item.start_time,
            item.therapist_name, source_cell
        )
        for item in slots
    )


def validate_weekly_entry_consistency(
    entries: tuple[SourceWeeklyEntry, ...] | list[SourceWeeklyEntry],
    *,
    confirmed_combined_slots: frozenset[frozenset[SourceWeeklyEntry]] = frozenset(),
) -> None:
    """Reject collisions unless the exact two-entry combined slot is confirmed.

    The caller must obtain explicit, source-bound operator confirmation.
    No implicit exception is granted by matching providers alone; stale
    confirmations and third appointments fail closed.
    """
    from .recurring_conflicts import _provider_key

    groups: dict[tuple[str, str, time], list[SourceWeeklyEntry]] = {}
    for entry in entries:
        groups.setdefault((entry.patient_id, entry.weekday, entry.start_time), []).append(entry)

    used = set()
    for group in groups.values():
        if len(group) == 1:
            continue
        pair = frozenset(group)
        providers = {_provider_key(item.therapist_name) for item in group}
        if (
            len(group) == len(pair) == 2
            and {item.treatment for item in group} == {"ΦΘ", "Ρομποτικό"}
            and len(providers) == 1 and "" not in providers
            and pair in confirmed_combined_slots
        ):
            used.add(pair)
            continue
        raise ProgrammeMappingError("Potential patient slot overlap without exact confirmation")
    if used != confirmed_combined_slots:
        raise ProgrammeMappingError("Stale or unmatched combined slot confirmation")
