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
