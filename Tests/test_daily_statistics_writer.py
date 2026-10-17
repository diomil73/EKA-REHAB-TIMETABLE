from __future__ import annotations

from datetime import date

from rehab_core.daily_statistics import DailyTreatmentStatistics
from rehab_excel.daily_statistics_writer import build_daily_statistics_rows


def test_daily_statistics_rows_surface_suspension_and_finalization():
    statistics = DailyTreatmentStatistics(
        scheduled_sessions=12,
        performed_sessions=8,
        replaced_sessions=2,
        patient_absent_sessions=1,
        cancelled_sessions=3,
        department_postponed_sessions=1,
        patient_no_show_sessions=1,
        treatment_suspended_sessions=1,
        unresolved_sessions=0,
    )

    rows = build_daily_statistics_rows(date(2026, 10, 19), statistics)
    values = {row.key: row.value for row in rows}

    assert values["date"] == "19/10/2026"
    assert values["scheduled"] == 12
    assert values["performed"] == 8
    assert values["replaced"] == 2
    assert values["treatment_suspended"] == 1
    assert values["unresolved"] == 0
    assert values["finalized"] == "ΝΑΙ"
