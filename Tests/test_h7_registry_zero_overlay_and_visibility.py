from datetime import date

import pytest

from rehab_excel.patient_history_reports import (
    PatientHistoryReportError,
    PatientHistoryReportRow,
    include_registered_patients_without_sessions,
)
from tools import export_patient_history_ui as exporter
from tools.inspect_h7_patient_report_visibility import _is_numeric_id


def _report(patient: str, start: str, end: str | None = None, sessions: int = 1):
    beginning = date.fromisoformat(start)
    return PatientHistoryReportRow(
        period_start=beginning,
        period_end=date.fromisoformat(end or start),
        patient_id=patient,
        finalized_days=1,
        scheduled=sessions,
        performed=sessions,
        patient_absences=0,
        cancellations=0,
        replacements=0,
    )


def test_daily_includes_registry_zero_rows_without_changing_history():
    history = (_report("1", "2026-10-30", sessions=2), _report("18", "2026-10-30"))
    rows = include_registered_patients_without_sessions(
        history, ("01", "15", "18", "19", "20", "21", "22"),
        mode="daily",
    )
    assert [row.patient_id for row in rows] == ["1", "15", "18", "19", "20", "21", "22"]
    assert len(rows) == 7
    assert sum(row.scheduled for row in rows) == 3
    assert rows[1].scheduled == rows[1].performed == rows[1].finalized_days == 0
    assert rows[0] is history[0]


def test_overall_multi_day_includes_one_zero_row_per_registry_member():
    history = (
        _report("1", "2026-10-29", "2026-10-30"),
        _report("2", "2026-10-30", "2026-11-04"),
    )
    rows = include_registered_patients_without_sessions(
        history, ("1", "2", "15"), mode="overall"
    )
    assert len(rows) == 3
    zero = next(row for row in rows if row.patient_id == "15")
    assert (zero.period_start, zero.period_end) == (date(2026, 10, 29), date(2026, 11, 4))
    assert zero.scheduled == 0


def test_weekly_two_periods_fill_per_period_not_globally():
    history = (
        _report("1", "2026-10-26", "2026-11-01"),
        _report("2", "2026-11-02", "2026-11-08"),
    )
    rows = include_registered_patients_without_sessions(
        history, ("1", "2"), mode="weekly"
    )
    assert len(rows) == 4
    assert sorted(row.scheduled for row in rows) == [0, 0, 1, 1]


def test_empty_finalized_history_does_not_invent_reporting_dates():
    assert include_registered_patients_without_sessions(
        (), ("1", "2"), mode="monthly"
    ) == ()


def test_duplicate_registry_and_unknown_history_fail_closed():
    with pytest.raises(PatientHistoryReportError, match="Duplicate patient ID"):
        include_registered_patients_without_sessions(
            (_report("1", "2026-10-30"),), ("01", "1"), mode="daily"
        )
    with pytest.raises(PatientHistoryReportError, match="absent from current registry"):
        include_registered_patients_without_sessions(
            (_report("99", "2026-10-30"),), ("1",), mode="daily"
        )


def test_export_opt_in_retains_default_history_only(monkeypatch, tmp_path):
    monkeypatch.setattr(
        exporter,
        "read_patient_history_report",
        lambda *args, **kwargs: (_report("1", "2026-10-30"),),
    )
    monkeypatch.setattr(
        exporter,
        "read_patient_registry",
        lambda *args, **kwargs: [
            type("Patient", (), {"patient_id": "1", "display_name": "SYNTHETIC ONE"})(),
            type("Patient", (), {"patient_id": "15", "display_name": "SYNTHETIC FIFTEEN"})(),
        ],
    )
    default = tmp_path / "history.tsv"
    exporter.export_patient_history_ui("unused.xlsm", default, mode="daily")
    assert len(default.read_text(encoding="utf-8-sig").splitlines()) == 2

    complete = tmp_path / "registry.tsv"
    exporter.export_patient_history_ui(
        "unused.xlsm", complete, mode="daily", include_registered_zero=True
    )
    lines = complete.read_text(encoding="utf-8-sig").splitlines()
    assert len(lines) == 3
    assert lines[2].split("\t") == [
        "2026-10-30", "2026-10-30", "15", "SYNTHETIC FIFTEEN",
        "0", "0", "0", "0", "0", "0", ""
    ]


def test_excel_ids_match_padded_and_unpadded_numeric_values():
    assert _is_numeric_id("018", 18)
    assert _is_numeric_id(22, 22)
    assert not _is_numeric_id("P18", 18)


def test_vba_expressly_labels_current_registry_and_zero_activity():
    from rehab_excel.patient_history_report_vba import PATIENT_REPORT_VBA

    assert '--include-registered-zero' in PATIENT_REPORT_VBA
    assert 'Τρέχον μητρώο: χωρίς συνεδρία = 0' in PATIENT_REPORT_VBA


def test_patient_report_rows_have_readable_height():
    from rehab_excel.patient_history_report_vba import PATIENT_REPORT_VBA

    assert 'If lastRow >= 6 Then ws.Rows("6:" & CStr(lastRow)).RowHeight = 21' in PATIENT_REPORT_VBA
