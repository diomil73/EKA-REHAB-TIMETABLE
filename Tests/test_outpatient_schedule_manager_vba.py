from rehab_excel.outpatient_schedule_edit_vba import EDIT_MODULE_CODE
from rehab_excel.outpatient_schedule_manager_vba import (
    FORM_CODE,
    MACRO_NAME,
    MODULE_CODE,
    patch_registration_menu_for_outpatient_manager,
)
from rehab_excel.registration_menu_vba import USERFORM_CODE


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
    assert 'lblPatient.Caption = "Ασθενής"' in FORM_CODE
    assert 'lblTreatment.Caption = "Θεραπεία"' in FORM_CODE
    assert 'lblTime.Caption = "Ώρα"' in FORM_CODE
    assert 'lblDays.Caption = "Ημέρες"' in FORM_CODE
    assert 'lblTherapist.Caption = "Θεραπευτής"' in FORM_CODE
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


def test_registration_menu_routes_outpatient_button_to_manager():
    patched = patch_registration_menu_for_outpatient_manager(USERFORM_CODE)

    assert 'StyleMenuButton cmdOutpatientSchedule, "Προγράμματα εξωτερικών ασθενών", 203' in patched
    assert "ShowOutpatientScheduleManager" in patched
    assert "frmOutpatientSchedule.Show" not in patched
    assert "Πρόγραμμα εξωτερικού ασθενή δεν είναι εγκατεστημένη" not in patched


def test_legacy_registration_menu_without_outpatient_entry_gets_one():
    legacy = USERFORM_CODE.replace(
        '    StyleMenuButton cmdOutpatientSchedule, "Πρόγραμμα εξωτερικού ασθενή", 203\n',
        "",
    )
    start = legacy.index("Private Sub cmdOutpatientSchedule_Click()")
    end = legacy.index("\nEnd Sub", start) + len("\nEnd Sub")
    legacy = legacy[:start] + legacy[end:]
    legacy = legacy.replace(".Top = 260", ".Top = 215")

    patched = patch_registration_menu_for_outpatient_manager(legacy)

    assert patched.count("StyleMenuButton cmdOutpatientSchedule") == 1
    assert '"Προγράμματα εξωτερικών ασθενών", 203' in patched
    assert patched.count("Private Sub cmdOutpatientSchedule_Click()") == 1
    assert "ShowOutpatientScheduleManager" in patched
    assert ".Top = 260" in patched
