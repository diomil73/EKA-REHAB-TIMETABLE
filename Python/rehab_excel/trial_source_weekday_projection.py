"""Pure A:V source-row to recurring-weekday candidates for the Diomil demo.

No source workbook access, imports, admissions, identities inferred from names,
native Excel writes, or clinical assertions. Only independently confirmed IDs
may be supplied by an upstream private admission process.

A returned candidate is NOT an approved schedule. Any blocked source field
blocks the entire row rather than silently dropping a treatment.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Sequence

from rehab_core.programme_mapping import (
    ProgrammeMappingError, SourceWeeklyEntry, make_weekly_entries,
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
    *, source_row: int, columns: Sequence[Any], confirmed_patient_id: str
) -> TrialSourceWeekdayPreview:
    """Expand only supported A:V weekday slots, without creating appointments.

    Blank weekdays with a populated treatment clock mean Monday-Friday for
    THIS source projection only. The general programme parser is unchanged.
    Blank clocks never generate sessions, even when weekdays are present.
    Such incomplete cells are retained as safety blockers rather than lost.

    F (reclining, 30 min) and U/V (afternoon/pool) are deliberately held
    until their separate operational contracts are implemented. The caller
    must independently verify the confirmed ID, source row and source hash.
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
            blockers.append(reason)

    # Never produce a partial apparently importable schedule.
    return TrialSourceWeekdayPreview(
        source_row=source_row,
        entries=tuple(candidates) if not blockers else (),
        blockers=tuple(blockers),
    )
