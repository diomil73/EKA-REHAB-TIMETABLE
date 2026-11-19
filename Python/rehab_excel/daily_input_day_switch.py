"""Pure DAILY_INPUT date-switch plan for two open service dates.

Preserve only editable input grid values, not rendered appointments, capacity,
attendance, or replacement facts. Those continue to use existing mechanisms.
The caller supplies the literal grid and actual finalized-date inventory from
one authoritative workbook. No file I/O, native writes or date navigation.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from math import isfinite
from typing import Iterable


class DailyInputDaySwitchError(ValueError):
    """Unsafe or stale input-grid transition."""


CellValue = str | int | float | bool | date | datetime | time | None
Grid = tuple[tuple[CellValue, ...], ...]


def _valid_value(value: CellValue) -> bool:
    # Excel formulas/errors must be rejected by the worksheet reader before
    # capture. All strings in this model are literal, including comments;
    # native rehydration must not reinterpret them as formulas.
    if value is None or type(value) in (str, int, bool, date):
        return True
    if type(value) is float:
        return isfinite(value)
    if type(value) in (datetime, time):
        return value.tzinfo is None
    return False


def _validate_grid(grid: Grid, width: int) -> None:
    if (not isinstance(grid, tuple) or not grid
            or any(not isinstance(row, tuple) or len(row) != width
                   or any(not _valid_value(value) for value in row)
                   for row in grid)):
        raise DailyInputDaySwitchError("Invalid DAILY_INPUT literal grid shape or value")


@dataclass(frozen=True)
class DailyInputDayDraft:
    service_date: date
    therapist_rows: Grid
    patient_rows: Grid

    def __post_init__(self) -> None:
        if type(self.service_date) is not date:
            raise DailyInputDaySwitchError("Service date must be a date, not a timestamp")
        # Existing DAILY_INPUT schema: A:F therapist, A:E patient.
        _validate_grid(self.therapist_rows, 6)
        _validate_grid(self.patient_rows, 5)

    @property
    def layout(self) -> tuple[int, int]:
        return len(self.therapist_rows), len(self.patient_rows)

    def empty_for(self, service_date: date) -> DailyInputDayDraft:
        return DailyInputDayDraft(
            service_date,
            tuple((None,) * 6 for _ in self.therapist_rows),
            tuple((None,) * 5 for _ in self.patient_rows),
        )


@dataclass(frozen=True)
class DailyInputDaySwitchPlan:
    outgoing: DailyInputDayDraft
    incoming: DailyInputDayDraft
    saved_before: tuple[DailyInputDayDraft, ...]
    saved_after: tuple[DailyInputDayDraft, ...]


def plan_open_day_switch(
    *,
    current: DailyInputDayDraft,
    saved_drafts: Iterable[DailyInputDayDraft],
    expected_current_date: date,
    target_date: date,
    finalized_dates: Iterable[date],
) -> DailyInputDaySwitchPlan:
    """Capture the outgoing draft and select the incoming grid, atomically in memory.

    The source and target must both be OPEN. Historical read/amendment and
    transition away from a finalized view are separate, existing history
    routes. The caller must validate grid formulas, protection, source hash,
    actual layout and finalized status, then persist/reload as one transaction.
    This plan never finalizes a day and grants no scheduling authorization.
    """
    if not isinstance(current, DailyInputDayDraft):
        raise DailyInputDaySwitchError("Current DAILY_INPUT grid required")
    if type(expected_current_date) is not date or current.service_date != expected_current_date:
        raise DailyInputDaySwitchError("Current service date changed before switch")
    if type(target_date) is not date or target_date == current.service_date:
        raise DailyInputDaySwitchError("Switch requires a different explicit service date")
    finals = tuple(finalized_dates)
    if any(type(day) is not date for day in finals):
        raise DailyInputDaySwitchError("Invalid finalized-date inventory")
    if current.service_date in finals or target_date in finals:
        raise DailyInputDaySwitchError("Finalized day requires the historical route")

    before = tuple(saved_drafts)
    by_date = {}
    for draft in before:
        if not isinstance(draft, DailyInputDayDraft):
            raise DailyInputDaySwitchError("Invalid stored DAILY_INPUT draft")
        if draft.service_date in by_date:
            raise DailyInputDaySwitchError("Duplicate stored service date")
        if draft.layout != current.layout:
            raise DailyInputDaySwitchError("Stored draft layout differs; no silent truncation")
        by_date[draft.service_date] = draft

    # Current live input is authoritative for the outgoing service date.
    by_date[current.service_date] = current
    incoming = by_date.get(target_date)
    if incoming is None:
        incoming = current.empty_for(target_date)
    after = tuple(by_date[day] for day in sorted(by_date))
    return DailyInputDaySwitchPlan(current, incoming, before, after)
