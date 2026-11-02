import inspect
from datetime import date

from rehab_core.therapist_history_reports import TherapistDailyReport
from rehab_excel.therapist_history_report_vba import (
    THERAPIST_HISTORY_MODULE_CODE,
    THERAPIST_HISTORY_MODULE_NAME,
    THERAPIST_HISTORY_NAV_BUTTON,
    THERAPIST_HISTORY_OPEN_MACRO,
    THERAPIST_HISTORY_SHEET_NAME,
    install_therapist_history_report,
)
from tools import export_therapist_history_ui as exporter
from tools import install_therapist_history_ui as installer


def _daily(therapist: str = "T1") -> TherapistDailyReport:
    return TherapistDailyReport(
        snapshot_id="finalized:2026-10-30:v2",
        service_date=date(2026, 10, 30),
        snapshot_version=2,
        therapist_id=therapist,
        capacity_timeslots=6,
        planned_occupied_timeslots=5,
        planned_sessions=6,
        final_occupied_timeslots=4,
        final_active_sessions=5,
        replacements_received=1,
        released_by_cancellation_or_no_show=1,
        treatment_suspensions_affecting_load=0,
        approved_shared_or_over_capacity_events=0,
    )


def test_ui_export_writes_stable_tsv_schema(monkeypatch, tmp_path):
    monkeypatch.setattr(exporter, "report_therapist_history", lambda *args, **kwargs: (_daily(),))
    output = tmp_path / "report.tsv"
    result = exporter.export_therapist_history_ui("ignored.xlsm", output, mode="daily")
    assert result["rows"] == 1
    lines = output.read_text(encoding="utf-8-sig").splitlines()
    assert lines[0].split("\t") == list(exporter.HEADERS)
    fields = lines[1].split("\t")
    assert fields[0:5] == ["2026-10-30", "2026-10-30", "2", "T1", "1"]
    assert fields[5:12] == ["5", "4", "6", "5", "1", "1", "0"]
    assert fields[12:] == ["83.3", "66.7"]


def test_ui_export_supports_weekly_monthly_and_overall(monkeypatch, tmp_path):
    monkeypatch.setattr(exporter, "report_therapist_history", lambda *args, **kwargs: (_daily(),))
    for mode in ("weekly", "monthly", "overall"):
        output = tmp_path / f"{mode}.tsv"
        result = exporter.export_therapist_history_ui("ignored.xlsm", output, mode=mode)
        assert result["mode"] == mode
        assert result["rows"] == 1
        assert len(output.read_text(encoding="utf-8-sig").splitlines()) == 2


def test_vba_report_runs_read_only_export_synchronously_in_same_excel_session():
    source = THERAPIST_HISTORY_MODULE_CODE
    assert THERAPIST_HISTORY_MODULE_NAME == "modTherapistHistoryReport"
    assert THERAPIST_HISTORY_SHEET_NAME == "THERAPIST_HISTORY_REPORT"
    assert THERAPIST_HISTORY_OPEN_MACRO == "OpenTherapistHistoryReport"
    assert "export_therapist_history_ui.py" in source
    assert 'CreateObject("WScript.Shell").Run(commandLine, 0, True)' in source
    assert 'RenderTherapistHistoryReport "daily"' in source
    assert 'RenderTherapistHistoryReport "weekly"' in source
    assert 'RenderTherapistHistoryReport "monthly"' in source
    assert 'RenderTherapistHistoryReport "overall"' in source
    assert "ThisWorkbook.Close" not in source
    assert "Application.Quit" not in source
    assert "Workbooks.Open" not in source


def test_user_facing_report_hides_snapshot_version_and_explains_each_mode():
    source = THERAPIST_HISTORY_MODULE_CODE
    assert 'ReportHeaders = Array("Ημερομηνία", "", "Θεραπευτής"' in source
    assert 'ReportHeaders = Array("Μήνας", "", "Θεραπευτής"' in source
    assert 'ReportHeaders = Array("Από", "Έως", "Θεραπευτής"' in source
    assert '"Έκδοση"' not in source
    assert '"Συνολικά για όλο το διαθέσιμο ιστορικό"' in source
    assert 'ws.Cells(rowIndex, 3).Value2 = parts(3)' in source
    assert 'parts(2)' not in source


def test_report_buttons_have_dedicated_visible_row_and_are_brought_to_front():
    source = THERAPIST_HISTORY_MODULE_CODE
    assert "ws.Rows(3).RowHeight = 32" in source
    assert "ws.Rows(4).RowHeight = 6" in source
    assert "ws.Rows(5).RowHeight = 30" in source
    assert "PositionReportButtons ws" in source
    assert "shape.Top = topPos" in source
    assert "shape.Height = 24" in source
    assert "shape.ZOrder 0" in source


def test_report_refresh_recreates_buttons_and_keeps_shapes_unprotected():
    source = THERAPIST_HISTORY_MODULE_CODE
    assert "EnsureReportButtons ws" in source
    assert 'EnsureOneReportButton ws, "ekaHistory_Daily"' in source
    assert 'EnsureOneReportButton ws, "ekaHistory_Weekly"' in source
    assert 'EnsureOneReportButton ws, "ekaHistory_Monthly"' in source
    assert 'EnsureOneReportButton ws, "ekaHistory_Overall"' in source
    assert 'EnsureOneReportButton ws, "ekaHistory_Back"' in source
    assert "shape.Visible = True" in source
    assert "ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True" in source


def test_each_mode_resets_window_to_top_and_refreezes_after_render():
    source = THERAPIST_HISTORY_MODULE_CODE
    assert "ResetReportWindow ws" in source
    assert "ActiveWindow.FreezePanes = False" in source
    assert "ActiveWindow.SplitColumn = 0" in source
    assert "ActiveWindow.SplitRow = 0" in source
    assert "ActiveWindow.ScrollColumn = 1" in source
    assert "ActiveWindow.ScrollRow = 1" in source
    assert 'ws.Range("A6").Select' in source
    assert "ActiveWindow.SplitRow = 5" in source
    assert "ActiveWindow.FreezePanes = True" in source


def test_report_ui_does_not_expose_or_mutate_technical_history_sheets():
    source = THERAPIST_HISTORY_MODULE_CODE
    assert "FINALIZED_DAY_SNAPSHOTS" not in source
    assert "THERAPIST_DAY_HISTORY" not in source
    assert "DAY_HISTORY_EVENTS" not in source
    assert 'Private Const REPORT_SHEET As String = "THERAPIST_HISTORY_REPORT"' in source


def test_installer_entry_is_on_therapist_daily_not_master():
    source = inspect.getsource(install_therapist_history_report)
    assert THERAPIST_HISTORY_NAV_BUTTON == "ekaToolbar_TherapistHistory"
    assert 'workbook.Worksheets("THERAPIST_DAILY")' in source
    assert 'workbook.Worksheets("MASTER_SCHEDULE")' not in source
    assert 'caption="Στατιστικά θεραπευτών"' in source
    assert "THERAPIST_HISTORY_REPORT" in source or "THERAPIST_HISTORY_SHEET_NAME" in source
    assert "report_ws.Unprotect()" in source


def test_native_installer_requires_technical_history_to_remain_very_hidden():
    source = inspect.getsource(installer.install_therapist_history_ui)
    assert '"FINALIZED_DAY_SNAPSHOTS"' in source
    assert '"THERAPIST_DAY_HISTORY"' in source
    assert '"DAY_HISTORY_EVENTS"' in source
    assert "if int(technical.Visible) != 2" in source
    assert "workbook.Save()" in source
