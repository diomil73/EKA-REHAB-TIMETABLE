"""Pure A:V source-row to recurring-weekday candidates for the Diomil demo.

No source workbook access, imports, admissions, identities inferred from names,
native Excel writes, or clinical assertions. Only independently confirmed IDs
may be supplied by an upstream private admission process.

A returned candidate is NOT an approved schedule. Any blocked source field
blocks the entire row rather than silently dropping a treatment.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Sequence, Mapping

from rehab_core.programme_mapping import (
    ProgrammeMappingError, SourceWeeklyEntry, make_weekly_entries, parse_clock, WEEKDAYS,
)


class TrialSourceWeekdayError(ValueError):
    """The proposed A:V row cannot be interpreted within the source contract."""


# (treatment, clock column, weekday column, optional provider column).
# Indices and column letters are intentionally explicit and source-bound.
SCHEDULE_FIELDS = (
    ("ΦΘ", "C", "D", "E"),
    ("Εργο", "G", "H", "I"),
    ("Λογο", "J", "K", "L"),
    ("Ψυχ", "M", "N", "O"),
    ("ΕΦΑ", "P", "Q", None),
    ("Ρομποτικό", "R", "S", "T"),
)
SOURCE_WIDTH = 22


@dataclass(frozen=True)
class TrialSourceWeekdayPreview:
    """Private candidate entries and privacy-safe blocking codes only."""

    source_row: int
    entries: tuple[SourceWeeklyEntry, ...] = field(repr=False)
    blockers: tuple[str, ...] = ()

    @property
    def ready_for_independent_review(self) -> bool:
        return bool(self.entries) and not self.blockers


def _present(value: Any) -> bool:
    return value is not None and bool(str(value).strip())


def _value(source: Sequence[Any], column: str) -> Any:
    return source[ord(column) - ord("A")]


def preview_source_weekdays(
    *, source_row: int, columns: Sequence[Any], confirmed_patient_id: str,
    historical_afternoon_acknowledged: bool = False,
    approved_morning_days: Mapping[str, tuple[str, ...]] | None = None,
    acknowledged_untimed_fields: frozenset[str] = frozenset(),
) -> TrialSourceWeekdayPreview:
    """Expand only supported A:V weekday slots, without creating appointments.

    Blank weekdays with a populated treatment clock mean Monday-Friday for
    THIS source projection only. The general programme parser is unchanged.
    Blank clocks never generate sessions, even when weekdays are present.
    Such incomplete cells are retained as safety blockers rather than lost.

    F/V require source-bound weekday and service approval; U is never a
    session. The caller verifies the identity and source hash.
    """
    if (
        not isinstance(source_row, int) or isinstance(source_row, bool)
        or source_row < 2 or len(columns) != SOURCE_WIDTH
        or not isinstance(confirmed_patient_id, str)
        or not confirmed_patient_id.strip() or not _present(columns[1])
    ):
        raise TrialSourceWeekdayError("Invalid source row shape or confirmed identity")

    candidates: list[SourceWeeklyEntry] = []
    blockers: list[str] = []
    for treatment, clock_col, day_col, provider_col in SCHEDULE_FIELDS:
        clock = _value(columns, clock_col)
        days = _value(columns, day_col)
        provider = _value(columns, provider_col) if provider_col else None
        address = f"{clock_col}{source_row}"
        if address in acknowledged_untimed_fields:
            if ((clock_col == "G" and clock == "*")
                    or (clock_col == "M" and not _present(clock)
                        and (_present(days) or _present(provider)))):
                continue  # No timed slot; original and private decision ledger retained.
            blockers.append(f"INVALID_UNTIMED_DISPOSITION_{clock_col}")
            continue
        if not _present(clock):
            if _present(days) or _present(provider):
                blockers.append(f"MISSING_CLOCK_{clock_col}")
            continue
        try:
            candidates.extend(make_weekly_entries(
                patient_id=confirmed_patient_id,
                treatment=treatment,
                source_cell=f"{clock_col}{source_row}",
                clock=clock,
                weekdays=days,
                therapist=provider,
                implicit_daily=not _present(days),
            ))
        except (ProgrammeMappingError, TypeError, ValueError):
            blockers.append(f"UNPARSEABLE_TREATMENT_{clock_col}")

    for column, reason in (
        ("F", "RECLINING_REQUIRES_SEPARATE_30_MIN_CONTRACT"),
        ("U", "AFTERNOON_UNMAPPED"),
        ("V", "POOL_UNMAPPED"),
    ):
        if _present(_value(columns, column)):
            if column == "U" and historical_afternoon_acknowledged is True:
                continue  # Explicitly excluded; never converted into a session.
            if column in {"F", "V"}:
                address = f"{column}{source_row}"
                days = (approved_morning_days or {}).get(address)
                if (isinstance(days, tuple) and days and len(set(days)) == len(days)
                        and all(day in WEEKDAYS for day in days)):
                    try:
                        clock = parse_clock(_value(columns, column))
                        if column == "F" and clock.hour * 60 + clock.minute + 30 > 1440:
                            raise ValueError("Reclining interval exceeds day")
                        candidates.extend(SourceWeeklyEntry(
                            confirmed_patient_id, "Ανακλινόμενο" if column == "F" else "Πισίνα",
                            day, clock, None, address,
                        ) for day in days)
                        continue
                    except (ProgrammeMappingError, TypeError, ValueError):
                        pass
            blockers.append(reason)

    # Do not infer durations of other treatments; authorized review must
    # establish full interval compatibility. The known reclining interval
    # must not contain any other start time on that patient's same day.
    for entry in candidates:
        if entry.treatment != "Ανακλινόμενο":
            continue
        start = entry.start_time.hour * 60 + entry.start_time.minute
        if any(other is not entry and other.weekday == entry.weekday
               and start <= other.start_time.hour * 60 + other.start_time.minute < start + 30
               for other in candidates):
            blockers.append("RECLINING_INTERVAL_CONFLICT")

    # Never produce a partial apparently importable schedule.
    return TrialSourceWeekdayPreview(
        source_row=source_row,
        entries=tuple(candidates) if not blockers else (),
        blockers=tuple(blockers),
    )
