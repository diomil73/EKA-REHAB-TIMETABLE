from datetime import date

import pytest

from rehab_excel.historical_statistics import DailyStatistics
from rehab_excel.period_statistics import group_daily_statistics, period_statistics_dicts


def day(value, scheduled=3, performed=2):
    return DailyStatistics(
        service_date=date.fromisoformat(value),
        snapshot_id=f"finalized:{value}:v2", version=2,
        scheduled=scheduled, performed=performed,
        patient_absences=scheduled-performed, cancellations=0,
        replacements=1 if performed else 0, therapists=2,
    )


def test_iso_week_crosses_year_boundary_correctly():
    rows = group_daily_statistics((day("2026-12-31"), day("2027-01-01")), period="week")
    assert len(rows) == 1
    result = rows[0]
    assert result.period == "2026-W53"
    assert result.start_date == date(2026, 12, 28)
    assert result.end_date == date(2027, 1, 3)
    assert (result.finalized_days, result.scheduled, result.performed) == (2, 6, 4)


def test_monthly_totals_distinguish_months():
    rows = group_daily_statistics((day("2026-10-30"), day("2026-10-31"), day("2026-11-02")), period="month")
    assert [(x.period, x.finalized_days, x.scheduled) for x in rows] == [
        ("2026-10", 2, 6), ("2026-11", 1, 3)
    ]


def test_daily_empty_and_serialization():
    assert group_daily_statistics((), period="day") == ()
    rows = group_daily_statistics((day("2026-10-30"),), period="day")
    serial = period_statistics_dicts(rows)
    assert serial[0]["start_date"] == "2026-10-30"
    assert serial[0]["end_date"] == "2026-10-30"
    assert serial[0]["replacements"] == 1


def test_duplicate_day_and_invalid_period_fail_closed():
    with pytest.raises(ValueError, match="Duplicate CURRENT_FINAL"):
        group_daily_statistics((day("2026-10-30"), day("2026-10-30")), period="month")
    with pytest.raises(ValueError, match="period must"):
        group_daily_statistics((), period="year")
