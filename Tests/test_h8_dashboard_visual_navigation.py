from rehab_excel.patient_history_report_vba import PATIENT_REPORT_VBA
from rehab_excel.therapist_history_report_vba import THERAPIST_HISTORY_MODULE_CODE


def test_mode_highlight_persists_in_both_reports_after_buttons_are_recreated():
    for code, name in (
        (PATIENT_REPORT_VBA, "HighlightPatientReportMode"),
        (THERAPIST_HISTORY_MODULE_CODE, "HighlightTherapistReportMode"),
    ):
        assert f"Private Sub {name}" in code
        assert f"{name} ws, reportMode" in code
        if name == "HighlightPatientReportMode":
            assert 'modes = Array("Daily", "Weekly", "Monthly", "Overall")' in code
        else:
            assert 'modes = Array("Today", "Yesterday", "Week", "Month", "Year", "Overall")' in code
        assert "RGB(37, 99, 170)" in code
        assert "RGB(255, 255, 255)" in code
        assert "RGB(236, 241, 247)" in code


def test_patient_chart_is_frozen_with_top_dashboard_when_scrolling_to_last_patient():
    src = PATIENT_REPORT_VBA
    assert 'ws.Range("A23").Select' in src
    assert "ActiveWindow.SplitRow = 22" in src
    assert "ActiveWindow.FreezePanes = True" in src
    assert 'ws.ChartObjects("ekaPatientOverviewChart").Delete' in src
    assert 'rowIndex = 23' in src


def test_therapists_show_three_native_panels_in_single_row():
    src = THERAPIST_HISTORY_MODULE_CODE
    for chart in ("ekaTherapistLoadChart", "ekaTherapistSessionChart", "ekaTherapistAttendanceChart"):
        assert f'ws.ChartObjects("{chart}").Delete' in src
    for segment in (('"A", "D"'), ('"E", "H"'), ('"I", "M"')):
        assert segment in src
    assert 'chartWidth - 4, 88)' in src
    assert 'chartWidth - 4, 108)' in src
    assert 'NewTherapistTrendChart' in src
    assert "plannedSessions, actualSessions" in src
    assert "plot.Chart.Axes(xlValue).MinimumScale = 0" in src
    assert "cohortAxisPatients = 6" in src


def test_workload_traffic_light_contract_is_explicit_and_never_looks_up_from_live_projection():
    src = THERAPIST_HISTORY_MODULE_CODE
    for threshold in ("loadPerDay >= 5.5", "loadPerDay >= 4.5", "loadPerDay >= 2.5", "loadPerDay > 0"):
        assert threshold in src
    assert "RGB(255, 3, 3)" in src
    assert "RGB(255, 180, 0)" in src
    assert "RGB(255, 242, 0)" in src
    assert "RGB(91, 250, 116)" in src
    assert "RGB(63, 168, 245)" in src
    assert "loadPerDay = uniquePatientsPerDay" in src
    assert 'ws.Cells(rowIndex, 14).Value2' in src


def test_absence_panel_marks_data_unavailable_instead_of_fabricating_zeros():
    src = THERAPIST_HISTORY_MODULE_CODE
    assert "Μη διαθέσιμο ιστορικό απουσιών / εργασίμων" in src
    assert 'plot.Chart.ChartTitle.Text = "Παρουσίες"' in src
    assert "If finalizedDays <= 0 Then" in src
    assert 'ws.Range("I19").Value2 = "Μη διαθέσιμο ιστορικό απουσιών / εργασίμων"' in src


def test_patient_three_categories_are_planned_performed_absent_only():
    src = PATIENT_REPORT_VBA
    assert 'series.XValues = Array("Προγραμματισμένες", "Πραγματοποιημένες", "Απουσίες")' in src
    assert "series.Values = Array(scheduled, delivered, absent)" in src
    assert "totalScheduled = totalScheduled + CLng(fields(5))" in src
    assert "BuildPatientOverviewChart ws, totalScheduled, totalDelivered, totalAbsent" in src


def test_therapist_load_title_omits_numeric_scale_without_removing_axis():
    src = THERAPIST_HISTORY_MODULE_CODE
    assert 'plot.Chart.ChartTitle.Text = "Φόρτος" & IIf(' in src
    assert '"Ασθενείς / ημέρα (0-6)"' not in src
    assert "plot.Chart.Axes(xlValue).MinimumScale = 0" in src
    assert "cohortAxisPatients = 6" in src
