from rehab_excel.outpatient_schedule_edit_vba import EDIT_MODULE_CODE
from rehab_excel.outpatient_schedule_manager_vba import (
    FORM_CODE,
    MACRO_NAME,
    MODULE_CODE,
)


def test_manager_is_the_business_facing_edit_entry_point():
    assert MACRO_NAME == "ShowOutpatientScheduleManager"
    assert "Public Sub ShowOutpatientScheduleManager()" in MODULE_CODE
    assert "frmOutpatientScheduleManager.Show" in MODULE_CODE
    assert "ShowOutpatientScheduleManager" in EDIT_MODULE_CODE
    assert "InputBox(" not in EDIT_MODULE_CODE


def test_manager_lists_business_fields_and_hides_internal_row_key():
    assert 'ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")' in FORM_CODE
    assert '.ColumnCount = 6' in FORM_CODE
    assert '.ColumnWidths = "0 pt;150 pt;95 pt;55 pt;85 pt;110 pt"' in FORM_CODE
    assert "patientName" in FORM_CODE
    assert "treatment" in FORM_CODE
    assert "timeText" in FORM_CODE
    assert "dayPattern" in FORM_CODE
    assert "therapistName" in FORM_CODE
    assert "lstSchedules.AddItem CStr(rowIndex)" in FORM_CODE


def test_manager_edit_selection_maps_hidden_row_to_internal_edit_action():
    assert "Private Function SelectedScheduleRow() As Long" in FORM_CODE
    assert "CLng(lstSchedules.List(lstSchedules.ListIndex, 0))" in FORM_CODE
    assert "EditOutpatientScheduleByRow targetRow" in FORM_CODE
    assert "Δώστε τη γραμμή" not in FORM_CODE
    assert "outpatient:" not in FORM_CODE


def test_manager_supports_create_edit_close_and_double_click():
    assert "Private Sub cmdNew_Click()" in FORM_CODE
    assert "ShowOutpatientScheduleForm" in FORM_CODE
    assert "Private Sub cmdEdit_Click()" in FORM_CODE
    assert "Private Sub lstSchedules_DblClick" in FORM_CODE
    assert "Private Sub cmdClose_Click()" in FORM_CODE
