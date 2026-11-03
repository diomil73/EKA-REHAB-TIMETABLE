from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date, timedelta
from pathlib import Path

from rehab_excel.historical_statistics import DailyStatistics, read_historical_statistics


@dataclass(frozen=True)
class PeriodStatistics:
    period: str
    start_date: date
    end_date: date
    finalized_days: int
    scheduled: int
    performed: int
    patient_absences: int
    cancellations: int
    replacements: int


def _period_bounds(day: date, period: str) -> tuple[str, date, date]:
    if period == "day":
        return day.isoformat(), day, day
    if period == "week":
        iso = day.isocalendar()
        first = day - timedelta(days=day.weekday())
        return f"{iso.year}-W{iso.week:02d}", first, first + timedelta(days=6)
    if period == "month":
        first = day.replace(day=1)
        if day.month == 12:
            following = date(day.year + 1, 1, 1)
        else:
            following = date(day.year, day.month + 1, 1)
        return f"{day.year}-{day.month:02d}", first, following - timedelta(days=1)
    raise ValueError("period must be day, week, or month")


def group_daily_statistics(
    days: tuple[DailyStatistics, ...],
    *,
    period: str,
) -> tuple[PeriodStatistics, ...]:
    """Aggregate verified finalized days, with no inferred empty-day sessions."""
    if period not in ("day", "week", "month"):
        raise ValueError("period must be day, week, or month")
    groups: dict[str, list[DailyStatistics]] = {}
    bounds: dict[str, tuple[date, date]] = {}
    seen_dates: set[date] = set()
    for item in days:
        if item.service_date in seen_dates:
            raise ValueError(f"Duplicate CURRENT_FINAL day: {item.service_date}")
        seen_dates.add(item.service_date)
        key, beginning, ending = _period_bounds(item.service_date, period)
        groups.setdefault(key, []).append(item)
        bounds[key] = beginning, ending

    return tuple(
        PeriodStatistics(
            period=key,
            start_date=bounds[key][0],
            end_date=bounds[key][1],
            finalized_days=len(group),
            scheduled=sum(d.scheduled for d in group),
            performed=sum(d.performed for d in group),
            patient_absences=sum(d.patient_absences for d in group),
            cancellations=sum(d.cancellations for d in group),
            replacements=sum(d.replacements for d in group),
        )
        for key, group in sorted(groups.items())
    )


def read_period_statistics(
    workbook_path: str | Path,
    *,
    period: str,
    start: date | None = None,
    end: date | None = None,
) -> tuple[PeriodStatistics, ...]:
    daily, _ = read_historical_statistics(workbook_path, start=start, end=end)
    return group_daily_statistics(daily, period=period)


def period_statistics_dicts(rows: tuple[PeriodStatistics, ...]) -> list[dict[str, object]]:
    """JSON-compatible representation, with explicit period boundaries."""
    return [
        {**asdict(row),
         "start_date": row.start_date.isoformat(),
         "end_date": row.end_date.isoformat()}
        for row in rows
    ]
