from rehab_excel.therapist_workload_strip_vba import (
    BUTTON_NAME,
    MODULE_CODE,
    MODULE_NAME,
    SHAPE_PREFIX,
)


def test_workload_strip_has_application_facing_refresh_action():
    assert MODULE_NAME == "modTherapistWorkloadStrip"
    assert BUTTON_NAME == "ekaWorkload_Refresh"
    assert SHAPE_PREFIX == "ekaWorkload_"
    assert "Public Sub ShowTherapistWorkload()" in MODULE_CODE
    assert 'ThisWorkbook.Worksheets("MASTER_SCHEDULE")' in MODULE_CODE
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY")' in MODULE_CODE
    assert '"Φόρτος θεραπευτών"' in MODULE_CODE


def test_workload_strip_counts_unique_occupied_timeslots_not_patient_lines():
    assert 'Set providers = CreateObject("Scripting.Dictionary")' in MODULE_CODE
    assert 'Set providerTimes = CreateObject("Scripting.Dictionary")' in MODULE_CODE
    assert "providerTimes(timeValue) = True" in MODULE_CODE
    assert "occupied = providerTimes.Count" in MODULE_CODE
    assert "DEFAULT_CAPACITY As Long = 6" in MODULE_CODE


def test_workload_strip_counts_only_operationally_active_lines():
    assert "Private Function CellHasActiveWork(ByVal cell As Range) As Boolean" in MODULE_CODE
    assert "If CellHasActiveWork(daily.Cells(rowIndex, col)) Then" in MODULE_CODE
    assert ".Font.Strikethrough" in MODULE_CODE
    assert ".Font.Italic" in MODULE_CODE
    assert 'If ch <> "→" Then' in MODULE_CODE
    assert "Len(Trim$(CStr(daily.Cells(rowIndex, col).Value))) > 0" not in MODULE_CODE


def test_workload_strip_merges_same_therapist_across_layout_blocks():
    assert "providerKey = NormalizeProviderName(providerName)" in MODULE_CODE
    assert "If Not providers.Exists(providerKey) Then" in MODULE_CODE
    assert "providerOrder.Add providerKey" in MODULE_CODE
    assert "Set providerTimes = providers(providerKey)" in MODULE_CODE
    assert "NormalizeProviderName" in MODULE_CODE


def test_workload_strip_uses_current_therapist_daily_layout_blocks():
    assert "headerRows = Array(1, 11)" in MODULE_CODE
    assert "firstRows = Array(2, 12)" in MODULE_CODE
    assert "lastRows = Array(8, 18)" in MODULE_CODE
    assert "For col = 2 To 10" in MODULE_CODE
    assert "TimeKey(daily.Cells(rowIndex, 1).Value)" in MODULE_CODE


def test_workload_strip_color_thresholds_are_green_orange_red():
    assert "If occupied >= DEFAULT_CAPACITY Then" in MODULE_CODE
    assert "ElseIf occupied = DEFAULT_CAPACITY - 1 Then" in MODULE_CODE
    assert "RGB(244, 176, 132)" in MODULE_CODE
    assert "RGB(255, 217, 102)" in MODULE_CODE
    assert "RGB(169, 208, 142)" in MODULE_CODE


def test_workload_strip_caption_shows_timeslot_capacity_fraction():
    assert 'providerName & " " & CStr(occupied) & "/" & CStr(DEFAULT_CAPACITY)' in MODULE_CODE
