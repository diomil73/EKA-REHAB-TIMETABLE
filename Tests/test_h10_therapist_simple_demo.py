from pathlib import Path

import pytest

from rehab_excel.therapist_history_report_vba import THERAPIST_HISTORY_MODULE_CODE
from tools.prepare_h10_therapist_simple_demo import (
    EVENT_CODE, H10DemoError, REQUIRED_VBA, _safe_module_lines, check_candidate, verify_selection_event,
)


def test_sticky_header_matches_patient_chart_behavior():
    assert "ActiveWindow.SplitRow = 22" in THERAPIST_HISTORY_MODULE_CODE
    assert 'ws.Range("A23").Select' in THERAPIST_HISTORY_MODULE_CODE
    assert "ActiveWindow.FreezePanes = True" in THERAPIST_HISTORY_MODULE_CODE


def test_selected_therapist_highlight_uses_one_active_row_and_resets_on_rerender():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'ws.Range("P1").Value2 = selectedRow' in vba
    assert 'ws.Range("P1").Value2 = 0' in vba
    assert 'Formula1:="=ROW()=$P$1"' in vba
    assert 'RGB(37, 99, 170)' in vba
    assert 'ws.Columns("P").Hidden = True' in vba
    assert 'Target.CountLarge <> 1' in EVENT_CODE
    assert "ShowSelectedTherapistChart" in EVENT_CODE


def test_period_benchmarks_are_separate_and_day_weighted_for_patients():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert "periodKey = CStr(ws.Cells(r, 1).Value2)" in vba
    assert "If selectedRow = 0 Or periodKey = selectedPeriod Then" in vba
    assert "referencePatients / referenceDays" in vba
    assert "referencePlanned / referenceCount" in vba
    assert "referenceActual / referenceCount" in vba
    assert '"Μ.Ο. περιόδου: "' in vba
    assert '"Κόκκινοι δείκτες: Μ.Ο. προγρ. "' in vba
    assert '" | πραγματ. "' in vba


def test_simple_labels_keep_technical_cause_out_of_primary_load_chart():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert vba.count('"Απουσίες/αναβολές"') == 3
    assert vba.count('"Αναστολές θεραπείας"') == 3
    assert vba.count('"Ασθενείς/ημέρα"') == 3
    assert '"Αποδεσμεύσεις"' not in vba
    assert '"Μη διαθέσιμο ιστορικό απουσιών / εργασίμων"' in vba
    assert "ws.Range(\"E19\").Value2" in vba


def test_h10_requires_strict_new_demo_destination_and_existing_bo(tmp_path):
    bo = tmp_path / "H8_DEMO_BO.xlsm"
    bo.write_bytes(b"sentinel")
    bq = tmp_path / "H10_DEMO_BU.xlsm"
    check_candidate(bo, bq)
    bq.write_bytes(b"already exists")
    with pytest.raises(H10DemoError, match="already exists"):
        check_candidate(bo, bq)
    assert bq.read_bytes() == b"already exists"
    with pytest.raises(H10DemoError, match="BU"):
        check_candidate(bo, tmp_path / "H10_DEMO_BV.xlsm")
    bn = tmp_path / "H8_DEMO_BN.xlsm"
    bn.write_bytes(b"frozen")
    with pytest.raises(H10DemoError, match="BO"):
        check_candidate(bn, tmp_path / "H10_DEMO_BU2.xlsm")


def test_install_requires_persisted_vba_markers_and_event_module():
    assert REQUIRED_VBA
    assert all(fragment in THERAPIST_HISTORY_MODULE_CODE for fragment in REQUIRED_VBA)
    assert EVENT_CODE.count("Worksheet_SelectionChange") == 1
    assert "ThisWorkbook.Save" in THERAPIST_HISTORY_MODULE_CODE


def test_empty_worksheet_vba_module_must_not_call_invalid_lines_one_zero():
    class EmptyModule:
        CountOfLines = 0

        def Lines(self, *args):
            raise AssertionError("Lines(1, 0) is invalid in native Excel")

    assert _safe_module_lines(EmptyModule()) == ""


def test_nonempty_vba_module_uses_valid_line_range():
    class Module:
        CountOfLines = 2

        def Lines(self, start, count):
            assert (start, count) == (1, 2)
            return "Option Explicit\\r\\nSub Test(): End Sub"

    assert _safe_module_lines(Module()).startswith("Option Explicit")


def test_bq_uses_real_vbproject_property_not_filename_replace_artifact():
    import inspect
    from tools import prepare_h10_therapist_simple_demo as module

    source = inspect.getsource(module)
    assert ".VBProject" in source
    assert "VBQroject" not in source
    assert "stage = " in source


def test_selection_event_contract_tolerates_vbide_crlf_and_spaces():
    variant = (
        "\r\nOption Explicit\r\n"
        + EVENT_CODE.replace("\n", "\r\n").replace(
            "    If ", "        If "
        ).replace(
            "    On Error", "       On   Error"
        )
    )
    verify_selection_event(variant)


def test_selection_event_contract_rejects_missing_guard():
    shortened = EVENT_CODE.replace(
        "    If Target.Row < 23 Then Exit Sub\n", ""
    )
    with pytest.raises(H10DemoError, match="statement readback"):
        verify_selection_event(shortened)


def test_selection_event_contract_rejects_duplicate_event():
    with pytest.raises(H10DemoError, match="declaration count 2"):
        verify_selection_event(EVENT_CODE + EVENT_CODE)


def test_selection_event_contract_rejects_absent_event():
    with pytest.raises(H10DemoError, match="declaration count 0"):
        verify_selection_event("Option Explicit\n")


def test_mean_markers_are_thin_short_chart_overlays_not_extra_bar_or_line_series():
    code = THERAPIST_HISTORY_MODULE_CODE
    assert "Private Sub AddTherapistMeanMarker" in code
    assert "plot.Chart.Shapes.AddLine(" in code
    assert "lineHalfWidth = slotWidth * 0.28" in code
    assert "marker.Line.Weight = 1.25" in code
    assert "marker.Line.DashStyle = msoLineSolid" in code
    assert "series.ChartType = xlLine" not in code
    assert "Μ.Ο. προγρ." in code
    assert "referencePatients, 1, 1, referenceColor" in code
    assert "referencePlanned, 1, 2, RGB(225, 0, 30)" in code
    assert "referenceActual, 2, 2, RGB(225, 0, 30)" in code


def test_mean_marker_changes_to_fluorescent_green_only_on_red_load_bar():
    code = THERAPIST_HISTORY_MODULE_CODE
    assert "referenceColor = RGB(225, 0, 30)" in code
    assert "If WorkloadTrafficColor(loadPerDay) = RGB(255, 3, 3) Then" in code
    assert "referenceColor = RGB(57, 255, 20)" in code


def test_department_totals_display_without_summing_therapist_percentages():
    code = THERAPIST_HISTORY_MODULE_CODE
    assert 'teamPlanned = CLng(referencePlanned)' in code
    assert 'teamActual = CLng(referenceActual)' in code
    assert 'ws.Range("A20").Value2 = "ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ | Προγραμματισμένες: "' in code
    assert 'CStr(teamPlanned)' in code
    assert 'CStr(teamActual)' in code
    assert 'If selectedRow = 0 Or periodKey = selectedPeriod Then' in code


def test_department_delivery_percent_is_count_weighted_and_zero_safe():
    code = THERAPIST_HISTORY_MODULE_CODE
    assert '" | Υλοποίηση: "' in code
    assert "100# * CDbl(teamActual) / Application.Max(1, teamPlanned)" in code
    assert 'Format$(100# * CDbl(teamActual)' in code
    assert '"0.0") & "%"' in code
    assert '"—"' in code
    assert round(359 / 434 * 100, 1) == 82.7


def test_all_therapists_share_cohort_scale_independent_of_selection():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'periodKey = CStr(ws.Cells(r, 1).Value2)' in vba
    assert 'If selectedRow = 0 Or periodKey = selectedPeriod Then' in vba
    assert 'cohortMaxPatients = Application.Max(cohortMaxPatients' in vba
    assert 'cohortMaxSessions = Application.Max(cohortMaxSessions' in vba
    assert 'CDbl(ws.Cells(r, 7).Value2), CDbl(ws.Cells(r, 8).Value2)' in vba
    assert 'plot.Chart.Axes(xlValue).MaximumScale = cohortAxisPatients' in vba
    assert 'plot.Chart.Axes(xlValue).MaximumScale = cohortAxisSessions' in vba
    assert 'Application.Max(plannedSessions, actualSessions' not in vba


def test_five_exact_user_hex_colors_are_trial_only():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'TRIAL #FF0303' in vba
    assert 'TRIAL #FFB400' in vba
    assert 'TRIAL #FFF200' in vba
    assert 'TRIAL #5BFA74' in vba
    assert 'TRIAL #3FA8F5' in vba
    for rgb in ['RGB(255, 3, 3)', 'RGB(255, 180, 0)',
                'RGB(255, 242, 0)', 'RGB(91, 250, 116)',
                'RGB(63, 168, 245)']:
        assert rgb in vba
    assert 'series.Points(1).Format.Fill.ForeColor.RGB = RGB(63, 168, 245)' in vba
    assert 'series.Points(2).Format.Fill.ForeColor.RGB = RGB(255, 3, 3)' in vba
    assert 'BLUE = planned, RED = delivered' in vba


def test_mini_trends_use_persisted_finalized_daily_cache_and_exact_colors():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert '" --trend-output " & QuoteReportArg(trendPath)' in vba
    assert 'ws.Range("Q1:U" & CStr(Application.Max(2' in vba
    assert 'CStr(trendLines(0)) <> _' in vba
    assert 'If UBound(trendParts) <> 4 Then' in vba
    assert 'ws.Cells(trendRow, 19).Value2 = CLng(trendParts(2))' in vba
    assert 'ws.Cells(trendRow, 20).Value2 = CLng(trendParts(3))' in vba
    assert 'ws.Columns("P:U").Hidden = True' in vba
    assert 'Private Sub BuildTherapistMiniTrends' in vba
    assert 'BuildTherapistMiniTrends ws, reportMode, selectedRow, therapistName' in vba
    assert '"ekaTherapistLoadTrend"' in vba
    assert '"ekaTherapistSessionTrend"' in vba
    assert 'AddTherapistTrendSeries plot, "Προγραμματισμένες", xPlot, pPlot, RGB(63, 168, 245)' in vba
    assert 'AddTherapistTrendSeries plot, "Πραγματοποιημένες", xPlot, dPlot, RGB(255, 3, 3)' in vba
    assert 'plot.Chart.Axes(xlValue).MaximumScale = Application.Max(1' in vba
    assert 'maxSession = Application.Max(maxSession, CDbl(ws.Cells(r, 19).Value2)' in vba
    assert 'If n < 2 Then Exit Sub' in vba
    assert 'ws.Range("E18").Value2 = "Μπλε: προγρ. | Κόκκινο: πραγματ."' in vba


def test_trend_window_modes_are_explicit_and_never_invent_attendance():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'Private Function TrendDateWithinReport' in vba
    assert 'Case "daily"' in vba
    assert 'Case "weekly"' in vba
    assert 'Case "monthly"' in vba
    assert 'Case "overall"' in vba
    assert 'dateText <= periodStart' in vba
    assert 'dateText >= periodStart And _' in vba
    assert 'Left$(dateText, 7) = Left$(periodStart, 7)' in vba
    assert 'ws.ChartObjects("ekaTherapistSessionTrend").Delete' in vba
    assert 'ws.ChartObjects("ekaTherapistLoadTrend").Delete' in vba
    assert '"Μη διαθέσιμο ιστορικό απουσιών / εργασίμων"' in vba
