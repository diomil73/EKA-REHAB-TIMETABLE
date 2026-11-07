from datetime import date
from types import SimpleNamespace

import pytest

from rehab_core.daily_state import DailySessionStatus as Status
from rehab_excel.patient_history_reports import (
    PatientHistoryReportError,
    summarize_patient_history_report,
)
from tools import export_patient_history_ui as exporter


def _row(day: str, session: str, patient: str, status: Status):
    return SimpleNamespace(
        service_date=date.fromisoformat(day), session_id=session,
        patient_id=patient, status=status,
    )


def test_daily_patient_counts_and_distinct_days():
    rows = (
        _row("2026-10-30", "s1", "p1", Status.ACTIVE),
        _row("2026-10-30", "s2", "p1", Status.PATIENT_ABSENT),
        _row("2026-10-30", "s3", "p2", Status.CANCELLED),
        _row("2026-10-31", "s4", "p1", Status.REPLACED),
    )
    result = summarize_patient_history_report(rows, mode="daily")
    assert len(result) == 3
    assert (result[0].patient_id, result[0].scheduled, result[0].performed, result[0].patient_absences, result[0].finalized_days) == ("p1", 2, 1, 1, 1)
    assert (result[-1].performed, result[-1].replacements) == (1, 1)


def test_weekly_iso_boundary_and_monthly_split():
    rows = (
        _row("2026-12-31", "s1", "p1", Status.ACTIVE),
        _row("2027-01-01", "s2", "p1", Status.PATIENT_ABSENT),
    )
    weekly = summarize_patient_history_report(rows, mode="weekly")
    assert len(weekly) == 1
    assert weekly[0].period_start == date(2026, 12, 28)
    assert weekly[0].period_end == date(2027, 1, 3)
    assert weekly[0].finalized_days == 2
    monthly = summarize_patient_history_report(rows, mode="monthly")
    assert len(monthly) == 2


def test_overall_dates_are_actual_and_filter_is_inclusive():
    rows = (
        _row("2026-10-29", "s1", "p1", Status.ACTIVE),
        _row("2026-10-30", "s2", "p1", Status.PATIENT_ABSENT),
    )
    result = summarize_patient_history_report(rows, mode="overall", start=date(2026, 10, 30))
    assert len(result) == 1
    assert result[0].period_start == result[0].period_end == date(2026, 10, 30)
    assert result[0].scheduled == 1


def test_duplicate_and_unresolved_fail_closed():
    row = _row("2026-10-30", "s1", "p1", Status.ACTIVE)
    with pytest.raises(PatientHistoryReportError, match="Duplicate"):
        summarize_patient_history_report((row, row), mode="monthly")
    with pytest.raises(PatientHistoryReportError, match="Unresolved"):
        summarize_patient_history_report(
            (_row("2026-10-30", "s1", "p1", Status.THERAPIST_ABSENT),),
            mode="daily",
        )


def test_invalid_mode_date_and_blank_patient():
    with pytest.raises(PatientHistoryReportError, match="Unsupported"):
        summarize_patient_history_report((), mode="yearly")
    with pytest.raises(PatientHistoryReportError, match="start date"):
        summarize_patient_history_report((), mode="daily", start=date(2026, 11, 1), end=date(2026, 10, 1))
    with pytest.raises(PatientHistoryReportError, match="Blank patient"):
        summarize_patient_history_report((_row("2026-10-30", "s1", " ", Status.ACTIVE),), mode="daily")


def test_export_stable_tsv_and_refuses_overwrite(monkeypatch, tmp_path):
    rows = (_row("2026-10-30", "s1", "p1", Status.ACTIVE),)
    result = summarize_patient_history_report(rows, mode="monthly")
    monkeypatch.setattr(exporter, "read_patient_history_report", lambda *a, **kw: result)
    monkeypatch.setattr(exporter, "read_patient_registry", lambda *a, **kw: (SimpleNamespace(patient_id="p1", display_name="ΜΑΡΙΑ"),))
    output = tmp_path / "patients.tsv"
    proof = exporter.export_patient_history_ui("ignored.xlsm", output, mode="monthly")
    assert proof["rows"] == 1
    lines = output.read_text(encoding="utf-8-sig").splitlines()
    assert lines[0].split("\t") == list(exporter.HEADERS)
    assert lines[1].split("\t") == ["2026-10-01", "2026-10-31", "p1", "ΜΑΡΙΑ", "1", "1", "1", "0", "0", "0", "100.0"]
    with pytest.raises(ValueError, match="already exists"):
        exporter.export_patient_history_ui("ignored.xlsm", output, mode="monthly")


def test_export_rejects_unsafe_tsv_patient_value(monkeypatch, tmp_path):
    rows = summarize_patient_history_report(
        (_row("2026-10-30", "s1", "p\t1", Status.ACTIVE),), mode="daily"
    )
    monkeypatch.setattr(exporter, "read_patient_history_report", lambda *a, **kw: rows)
    monkeypatch.setattr(exporter, "read_patient_registry", lambda *a, **kw: ())
    with pytest.raises(ValueError, match="control character"):
        exporter.export_patient_history_ui("ignored.xlsm", tmp_path / "unsafe.tsv")
    assert not (tmp_path / "unsafe.tsv").exists()
