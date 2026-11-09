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
    assert 'ws.Columns("P:U").Hidden = True' in vba
    assert 'Target.CountLarge <> 1' in EVENT_CODE
    assert "ShowSelectedTherapistChart" in EVENT_CODE


def test_period_benchmarks_are_separate_and_day_weighted_for_patients():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert "periodKey = CStr(ws.Cells(r, 1).Value2)" in vba
    assert "If selectedRow = 0 Or periodKey = selectedPeriod Then" in vba
    assert "referencePatients / referenceDays" in vba
    assert "referencePlanned / referenceCount" in vba
    assert "referenceActual / referenceCount" in vba
    assert '"Μ.Ο.: "' in vba
    assert 'referencePatients, "0.0"' in vba
    assert '"Κόκκινοι δείκτες: Μ.Ο. προγρ. "' in vba
    assert '" | πραγματ. "' in vba


def test_simple_labels_keep_technical_cause_out_of_primary_load_chart():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert vba.count('"Απουσίες/αναβολές"') == 3
    assert vba.count('"Αναστολές θεραπείας"') == 3
    # Three report headers plus the deliberate mini-trend series label.
    assert vba.count('"Ασθενείς/ημέρα"') == 4
    assert '"Αποδεσμεύσεις"' not in vba
    assert '"Μη διαθέσιμο ιστορικό απουσιών / εργασίμων"' in vba
    assert "ws.Range(\"E19\").Value2" in vba


def test_h10_requires_corrected_ca_and_strict_fresh_bz_destination(tmp_path):
    ca = tmp_path / "H8_DEMO_HISTORY_CA.xlsm"
    ca.write_bytes(b"sentinel")
    bz = tmp_path / "H10_DEMO_BZ.xlsm"
    check_candidate(ca, bz)
    bz.write_bytes(b"already exists")
    with pytest.raises(H10DemoError, match="already exists"):
        check_candidate(ca, bz)
    assert bz.read_bytes() == b"already exists"
    with pytest.raises(H10DemoError, match="BZ"):
        check_candidate(ca, tmp_path / "H10_DEMO_BY.xlsm")
    with pytest.raises(H10DemoError, match="BZ"):
        check_candidate(ca, tmp_path / "H10_DEMO_BZ2.xlsm")
    bo = tmp_path / "H8_DEMO_BO.xlsm"
    bo.write_bytes(b"frozen")
    with pytest.raises(H10DemoError, match="CA"):
        check_candidate(bo, tmp_path / "H10_DEMO_NEW_BZ.xlsm")
    assert bo.read_bytes() == b"frozen"


def test_bz_installer_checks_corrected_source_before_copy():
    import inspect
    from tools import prepare_h10_therapist_simple_demo as installer
    source = inspect.getsource(installer.install_native)
    assert source.index("verify_demo_history(source)") < source.index("output.open(\"xb\")")
    assert source.index("capacity_anomalies(") < source.index("output.open(\"xb\")")
    assert source.index("clock_conflicts(") < source.index("output.open(\"xb\")")
    assert '"SHARED_CLOCK_WITHOUT_RECORDED_APPROVAL" in exception["flags"]' in source
    assert '"SLOT_OR_SESSION_RECONCILIATION_MISMATCH" in exception["flags"]' in source
    assert '"OCCUPIED_SLOTS_EXCEED_CAPACITY" in row["flags"]' in source
    assert '"CAPACITY_NOT_CONFIGURED" in row["flags"]' in source



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
    # The new mini-trends ARE true line charts. Only the mean overlay
    # must never be represented as another chart data series.
    assert 'series.Name = "Μ.Ο. περιόδου"' not in code
    assert 'series.Name = "Μ.Ο. προγρ."' not in code
    assert 'series.Name = "Μ.Ο. πραγμ."' not in code
    assert "Μ.Ο. προγρ." in code
    assert "referencePatients, 1, 1, referenceColor" in code
    assert "referencePlanned, 1, 2, RGB(225, 0, 30)" in code
    assert "referenceActual, 2, 2, referenceColor" in code


def test_mean_marker_changes_to_fluorescent_green_only_on_red_load_bar():
    code = THERAPIST_HISTORY_MODULE_CODE
    assert "referenceColor = RGB(225, 0, 30)" in code
    assert "If WorkloadTrafficColor(loadPerDay) = RGB(255, 3, 3) And _" in code
    assert "referencePatients <= loadPerDay Then" in code
    assert "If referenceActual <= actualSessions And actualSessions > 0 Then" in code
    assert "referenceColor = RGB(57, 255, 20)" in code


def test_department_totals_display_without_summing_therapist_percentages():
    code = THERAPIST_HISTORY_MODULE_CODE
    assert 'teamPlanned = teamPlanned + CLng(ws.Cells(r, 7).Value2)' in code
    assert 'teamActual = teamActual + CLng(ws.Cells(r, 8).Value2)' in code
    assert 'ws.Range("A20").Value2 = "ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ (προβολή) | Προγραμματισμένες: "' in code
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
    assert "cohortAxisPatients = 6" in vba
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
    assert 'AddTherapistTrendSeries plot, "Αρχικά προγραμματισμένες", xPlot, pPlot, RGB(63, 168, 245)' in vba
    assert 'AddTherapistTrendSeries plot, "Τελικά πραγματοποιημένες", xPlot, dPlot, RGB(255, 3, 3)' in vba
    assert 'plot.Chart.Axes(xlValue).MaximumScale = Application.Max(1' in vba
    assert 'maxSession = Application.Max(maxSession, CDbl(ws.Cells(r, 19).Value2)' in vba
    assert 'If n < 2 Then' in vba
    assert 'ws.Range("A18").Value2 = "Τάση: ανεπαρκείς ημέρες"' in vba
    assert 'StyleTherapistTrendRuler plot' in vba
    assert 'ws.Range("E18").Value2 = "Μπλε: αρχικά προγρ. | Κόκκινο: τελικά πραγματ."' in vba


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


def test_department_appendix_is_after_last_therapist_and_never_selectable_as_person():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'Private Sub BuildDepartmentAppendix' in vba
    assert 'BuildDepartmentAppendix ws, rowIndex - 1, reportMode' in vba
    assert 'headerRow = finalTherapistRow + 3' in vba
    assert 'groupKey = "Όλο το ιστορικό"' in vba
    assert 'chart.Name = "ekaDepartmentHistoryChart"' in vba
    assert 'allDelivered / allPlanned' in vba
    assert 'If allPlanned > 0 Then' in vba
    assert 'ws.Range("P3").Value2 = rowIndex - 1' in vba
    assert 'finalRow = CLng(ws.Range("P3").Value2)' in vba
    assert 'Target.Row > CLng(Me.Range("P3").Value2)' in EVENT_CODE
    assert 'If Target.Row > CLng(Me.Range("P3").Value2) Then Exit Sub' in EVENT_CODE


def test_visible_mini_axis_gridlines_and_actual_over_plan_warning():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert '.HasMajorGridlines = True' in vba
    assert 'plot.Chart.Axes(xlValue).MajorUnit = 1' in vba
    assert 'chartWidth - 4, 76)' in vba
    assert 'positiveDeltaDays = positiveDeltaDays + 1' in vba
    assert 'If positiveDeltaDays > 0 Then' in vba
    assert 'Αρχικά προγραμματισμένες' in vba
    assert 'Τελικά πραγματοποιημένες' in vba


def test_four_mode_internal_management_preview_never_claims_private_xlsm_is_shareable():
    vba = THERAPIST_HISTORY_MODULE_CODE
    for name in ("ShowManagementSummaryDaily", "ShowManagementSummaryWeekly",
                 "ShowManagementSummaryMonthly", "ShowManagementSummaryOverall",
                 "ExportManagementOnlyWorkbook", "RenderManagementSummary"):
        assert f"Sub {name}" in vba
    assert 'Private Const MANAGEMENT_SHEET As String = "MANAGEMENT_SUMMARY"' in vba
    assert 'ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True' in vba
    assert 'BuildDepartmentAppendix ws, rowIndex - 1, reportMode' in vba
    assert 'RenderManagementSummary ws, rowIndex - 1, reportMode' in vba
    assert 'Private Sub RenderManagementSummary' in vba
    assert 'Αρχικά προγραμματισμένες' in vba
    assert 'Τελικά πραγματοποιημένες' in vba
    assert 'ws.Cells(targetRow + 2, 1).Value2' in vba
    assert '«Εξαγωγή μόνο διοίκησης»' in vba
    assert 'If Len(Dir$(path)) > 0 Then' in vba
    assert 'ExportManagementOnlyWorkbook' in vba


def test_native_installer_reserves_unique_by_and_verifies_management_sheet():
    import inspect
    from tools import prepare_h10_therapist_simple_demo as installer
    source = inspect.getsource(installer)
    assert 'MANAGEMENT_SUMMARY' in source
    assert 'Management preview protection invalid' in source
    assert 'Protected' not in source or 'management.Protect' in source
    assert '_BZ.XLSM' in source


def test_therapist_selection_does_not_change_entire_department_total():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert "teamPlanned = teamPlanned + CLng(ws.Cells(r, 7).Value2)" in vba
    assert "teamActual = teamActual + CLng(ws.Cells(r, 8).Value2)" in vba
    assert "If selectedRow = 0 Or periodKey = selectedPeriod Then" in vba
    assert vba.index("teamPlanned = teamPlanned +") < vba.index(
        "If selectedRow = 0 Or periodKey = selectedPeriod Then"
    )
    assert "teamPlanned = CLng(referencePlanned)" not in vba
    assert "teamActual = CLng(referenceActual)" not in vba


def test_imported_vba_verification_reports_exact_missing_markers():
    import inspect
    from tools import prepare_h10_therapist_simple_demo as installer
    source = inspect.getsource(installer._verify_native)
    assert "missing = [" in source
    assert "H10 VBA marker read-back mismatch" in source
    assert "persisted lines=" in source
    assert "normalize(fragment) not in installed_normalized" in source


def test_management_protection_uses_explicit_com_arguments_and_reads_flags():
    import inspect
    from tools import prepare_h10_therapist_simple_demo as mod
    install = inspect.getsource(mod.install_native)
    audit = inspect.getsource(mod._verify_native)
    assert 'management.Protect("", False, True, True)' in install
    assert "managed_contents = bool(management.ProtectContents)" in audit
    assert "managed_drawing = bool(management.ProtectDrawingObjects)" in audit
    assert "managed_scenarios = bool(management.ProtectScenarios)" in audit
    assert "DrawingObjects={managed_drawing}" in audit


def test_report_refresh_clears_old_traffic_colors_and_department_counts_are_numeric():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'ws.Range("A22:N2000").ClearFormats' in vba
    assert 'ws.Range(ws.Cells(startRow, 2), _' in vba
    assert 'ws.Cells(lastGroup + 1, 3)).NumberFormat = "0"' in vba
    assert 'ws.Columns("B:C").NumberFormat = "0"' in vba


def test_individual_selection_does_not_delete_the_persistent_department_chart():
    vba = THERAPIST_HISTORY_MODULE_CODE
    start = vba.index("Private Sub BuildTherapistLoadChart(")
    end = vba.index("End Sub", start)
    per_therapist = vba[start:end]
    assert 'ws.ChartObjects("ekaDepartmentHistoryChart").Delete' not in per_therapist
    assert 'chart.Name = "ekaDepartmentHistoryChart"' in vba
    assert 'ws.ChartObjects("ekaDepartmentHistoryChart").Delete' in vba


def test_management_internal_dashboard_does_not_obscure_buttons_and_has_kpi_tiles():
    vba = THERAPIST_HISTORY_MODULE_CODE
    begin = vba.index("Private Sub RenderManagementSummary(")
    end = vba.index("Private Sub BuildDepartmentAppendix(", begin)
    source = vba[begin:end]
    assert 'ws.Range("A8:B8").Merge' in source
    assert 'ws.Range("C8:D8").Merge' in source
    assert 'ws.Range("F8:G8").Merge' in source
    assert 'ws.Range("A9").Value2 = plannedTotal' in source
    assert 'ws.Range("C9").Value2 = deliveredTotal' in source
    assert 'ws.Range("F9").Value2 = deliveredTotal / plannedTotal' in source
    assert 'ws.Range("F15").Left' in source
    assert 'ws.Range("F15").Top' in source
    assert 'shape.Top = ws.Range("A5").Top' in source
    assert 'shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB' in source
    assert 'ActiveWindow.DisplayHeadings = False' in source
    assert 'ActiveWindow.DisplayGridlines = False' in source
    assert 'ws.Range("A15:D15").Interior.Color' in source


def test_workload_charts_are_fixed_zero_to_six_and_do_not_hide_excess_facts():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert "cohortAxisPatients = 6" in vba
    assert "plot.Chart.Axes(xlValue).MaximumScale = cohortAxisPatients" in vba
    assert "plot.Chart.Axes(xlValue).MaximumScale = 6" in vba
    assert "plot.Chart.Axes(xlValue).MajorUnit = 1" in vba
    assert 'If loadPerDay > 6 Then' in vba
    assert 'If maxPatient > 6 Then' in vba
    assert 'Format$(loadPerDay, "0.0")' in vba
    # Actual count is preserved even when the fixed chart cannot show > 6.
    assert "loadPerDay = uniquePatientsPerDay" in vba
    assert "series.Values = Array(loadPerDay)" in vba


def test_compact_chart_titles_and_clear_vertical_separation():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'chartWidth - 4, 88)' in vba
    assert 'ws.Range(firstColumn & "11").Top' in vba
    assert 'chartWidth - 4, 76)' in vba
    assert vba.count('plot.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 9') == 3
    assert 'plot.Chart.ChartTitle.Text = "Φόρτος"' in vba
    assert 'plot.Chart.ChartTitle.Text = "Συνεδρίες"' in vba


def test_department_period_labels_use_displayed_dates_not_excel_serials():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'groupKey = CStr(ws.Cells(r, 1).Text)' in vba
    assert 'groupKey = groupKey & " έως " & CStr(ws.Cells(r, 2).Text)' in vba


def test_six_current_calendar_mode_buttons_are_distinct_and_old_shapes_retired():
    vba = THERAPIST_HISTORY_MODULE_CODE
    pairs = [
        ("Today", "Σήμερα", "ShowTherapistHistoryToday", "today"),
        ("Yesterday", "Χθες", "ShowTherapistHistoryYesterday", "yesterday"),
        ("Week", "Εβδομάδα", "ShowTherapistHistoryCurrentWeek", "week"),
        ("Month", "Μήνας", "ShowTherapistHistoryCurrentMonth", "month"),
        ("Year", "Έτος", "ShowTherapistHistoryCurrentYear", "year"),
        ("Overall", "Σύνολο", "ShowTherapistHistoryOverall", "overall"),
    ]
    for key, label, macro, mode in pairs:
        assert f'EnsureOneReportButton ws, "ekaHistory_{key}", "{label}", "{macro}"' in vba
        assert f'Case "{mode}": selected = "{key}"' in vba
    assert 'obsolete = Array("ekaHistory_Daily", "ekaHistory_Weekly", "ekaHistory_Monthly")' in vba
    assert 'ws.Shapes(CStr(oldName)).Delete' in vba
    assert 'RenderTherapistHistoryReport "today"' in vba
    assert 'Case "today", "yesterday", "week", "month", "year"' in vba
