from __future__ import annotations

from datetime import date
from types import SimpleNamespace

import pytest

from rehab_core.daily_state import DailySessionStatus
from tools import export_therapist_history_ui as mod


def daily(day: str, name: str, planned: int, actual: int):
    return SimpleNamespace(
        service_date=date.fromisoformat(day),
        therapist_id=name,
        snapshot_version=1,
        planned_occupied_timeslots=planned,
        final_occupied_timeslots=actual,
        planned_sessions=planned,
        final_active_sessions=actual,
        replacements_received=0,
        released_by_cancellation_or_no_show=planned - actual,
        treatment_suspensions_affecting_load=0,
        planned_capacity_utilization=None,
        actual_capacity_utilization=None,
    )


def session(day: str, name: str, patient: str, identifier: str):
    return SimpleNamespace(
        service_date=date.fromisoformat(day),
        effective_therapist_id=name,
        patient_id=patient,
        session_id=identifier,
        status=DailySessionStatus.ACTIVE,
    )


def test_h10_optional_daily_trend_is_sorted_and_does_not_change_report_tsv(tmp_path, monkeypatch):
    # Source deliberately out of date order; output MUST sort chronologically
    # before the VBA reads it. No invented dates or patient records.
    reports = (
        daily("2026-10-30", "T1", 5, 4),
        daily("2026-10-29", "T2", 7, 6),
        daily("2026-10-29", "T1", 4, 3),
    )
    monkeypatch.setattr(mod, "report_therapist_history", lambda _: reports)
    monkeypatch.setattr(
        mod, "read_current_patient_session_history_rows",
        lambda _: (
            session("2026-10-29", "T1", "P1", "S1"),
            session("2026-10-29", "T1", "P1", "S2"),
            session("2026-10-30", "T1", "P2", "S3"),
            session("2026-10-29", "T2", "P3", "S4"),
        ),
    )
    workbook = tmp_path / "synthetic.xlsm"
    report = tmp_path / "report.tsv"
    trend = tmp_path / "trend.tsv"

    result = mod.export_therapist_history_ui(
        workbook, report, mode="daily", trend_output_path=trend,
    )
    assert result["rows"] == 3
    main = report.read_text(encoding="utf-8-sig").splitlines()
    assert main[0].split("\t") == list(mod.HEADERS)
    assert all(len(line.split("\t")) == len(mod.HEADERS) for line in main[1:])

    lines = trend.read_text(encoding="utf-8-sig").splitlines()
    assert lines == [
        "Date\tTherapist\tPlanned\tDelivered\tUniquePatients",
        "2026-10-29\tT1\t4\t3\t1",
        "2026-10-29\tT2\t7\t6\t1",
        "2026-10-30\tT1\t5\t4\t1",
    ]


def test_h10_trend_output_is_optional_and_refuses_collision(tmp_path, monkeypatch):
    monkeypatch.setattr(mod, "report_therapist_history", lambda _: ())
    monkeypatch.setattr(mod, "read_current_patient_session_history_rows", lambda _: ())
    workbook = tmp_path / "synthetic.xlsm"
    dest = tmp_path / "report.tsv"
    result = mod.export_therapist_history_ui(workbook, dest, mode="daily")
    assert result["rows"] == 0
    assert dest.exists()
    dest.unlink()
    with pytest.raises(mod.TherapistHistoryUiExportError, match="collides"):
        mod.export_therapist_history_ui(
            workbook, dest, mode="daily", trend_output_path=dest,
        )
    assert not dest.exists()


@pytest.mark.parametrize(
    ("mode", "expected_start", "expected_end"),
    [
        ("today", "2026-11-09", "2026-11-09"),
        ("yesterday", "2026-11-08", "2026-11-08"),
        ("week", "2026-11-09", "2026-11-09"),
        ("month", "2026-11-01", "2026-11-09"),
        ("year", "2026-01-01", "2026-11-09"),
        ("overall", None, "2026-11-09"),
    ],
)
def test_calendar_mode_window_is_explicit_and_never_future_filled(mode, expected_start, expected_end):
    start, end = mod.selected_period(mode, date(2026, 11, 9))
    assert (start.isoformat() if start else None) == expected_start
    assert end.isoformat() == expected_end


def test_six_modes_emit_at_most_one_row_per_therapist_and_no_future_facts(tmp_path, monkeypatch):
    observations = (
        daily("2026-11-06", "T1", 2, 1),
        daily("2026-11-07", "T1", 3, 2),
        daily("2026-11-09", "T1", 4, 3),
        daily("2026-11-09", "T2", 5, 4),
        daily("2026-11-10", "T1", 6, 5),
    )
    monkeypatch.setattr(mod, "report_therapist_history", lambda _: observations)
    monkeypatch.setattr(mod, "read_current_patient_session_history_rows", lambda _: ())
    workbook = tmp_path / "demo.xlsm"
    expected = {"today": 2, "yesterday": 0, "week": 2, "month": 2, "year": 2, "overall": 2}
    for mode, count in expected.items():
        output = tmp_path / f"{mode}.tsv"
        trend = tmp_path / f"{mode}-trend.tsv"
        result = mod.export_therapist_history_ui(
            workbook, output, mode=mode, as_of=date(2026, 11, 9), trend_output_path=trend
        )
        assert result["rows"] == count
        body = output.read_text(encoding="utf-8-sig").splitlines()[1:]
        assert len({line.split("\t")[3] for line in body}) == len(body)
        assert all("2026-11-10" not in line for line in trend.read_text(encoding="utf-8-sig").splitlines())
    today = (tmp_path / "today.tsv").read_text(encoding="utf-8-sig").splitlines()
    assert today[1].split("\t")[7:9] == ["4", "3"]
    assert len((tmp_path / "yesterday.tsv").read_text(encoding="utf-8-sig").splitlines()) == 1
