from datetime import date
from types import SimpleNamespace

from rehab_core.daily_state import DailySessionStatus
from rehab_excel.patient_history_reports import summarize_patient_history_report
from rehab_excel.patient_history_report_vba import (
    PATIENT_REPORT_VBA,
    install_patient_report_components,
)
from tools.export_patient_history_ui import _display_patient_id
from tools import install_patient_history_ui


def _row(patient: str, session: str):
    return SimpleNamespace(
        patient_id=patient, session_id=session, service_date=date(2026, 10, 30),
        status=DailySessionStatus.ACTIVE,
    )


def test_numeric_patient_ids_naturally_ordered():
    values = ("20", "2", "10", "1", "3", "11")
    report = summarize_patient_history_report(
        tuple(_row(value, f"s{i}") for i, value in enumerate(values)), mode="monthly"
    )
    assert [r.patient_id for r in report] == ["1", "2", "3", "10", "11", "20"]


def test_patient_codes_padded_for_display_only():
    assert [_display_patient_id(v) for v in ("1", "2", "9", "10", "20")] == [
        "01", "02", "09", "10", "20"
    ]
    assert _display_patient_id("P1") == "P1"
    assert _display_patient_id("001") == "001"
    assert _display_patient_id("01") == "01"


def test_h7_entry_requires_h4_button_and_aligns_on_therapist_daily():
    import inspect

    src = inspect.getsource(install_patient_report_components)
    assert 'workbook.Worksheets("THERAPIST_DAILY")' in src
    assert 'source.Shapes("ekaToolbar_TherapistHistory")' in src
    assert 'source.Shapes(PATIENT_REPORT_NAV)' in src
    assert 'daily.Shapes(PATIENT_REPORT_NAV).Delete()' in src
    assert "float(therapist_button.Top)" in src
    assert 'shape.OnAction = PATIENT_REPORT_OPEN' in src


def test_h7_reopen_audit_checks_no_daily_entry_or_overlap():
    import inspect

    src = inspect.getsource(install_patient_history_ui._audit)
    assert 'workbook.Worksheets("THERAPIST_DAILY").Shapes(PATIENT_REPORT_NAV)' in src
    assert 'workbook.Worksheets("DAILY_INPUT").Shapes(PATIENT_REPORT_NAV)' in src
    assert "statistics buttons misaligned" in src
    assert "navigation buttons overlap" in src


def test_excel_patient_code_text_not_converted_to_number():
    assert 'ws.Columns("C").NumberFormat = "@"' in PATIENT_REPORT_VBA
