import inspect

from rehab_excel.application_shell_vba import (
    APP_SHELL_MODULE_CODE,
    MASTER_SIDEBAR_FORM_CODE,
    MASTER_SIDEBAR_FORM_NAME,
    _install_master_sidebar_form,
)


def test_master_sidebar_is_scroll_independent_modeless_form():
    assert MASTER_SIDEBAR_FORM_NAME == "frmMasterSidebar"
    assert "Public Sub ShowMasterSidebar()" in APP_SHELL_MODULE_CODE
    assert ".Show vbModeless" in APP_SHELL_MODULE_CODE
    assert ".Left = Application.Left + 4" in APP_SHELL_MODULE_CODE
    assert ".Top = Application.Top + 72" in APP_SHELL_MODULE_CODE
    assert "Application.CommandBars.Add" not in APP_SHELL_MODULE_CODE


def test_master_sidebar_contains_all_four_master_actions():
    assert "ToolbarRegistryManagement" in MASTER_SIDEBAR_FORM_CODE
    assert "ToolbarAbsences" in MASTER_SIDEBAR_FORM_CODE
    assert "ToolbarTherapistDaily" in MASTER_SIDEBAR_FORM_CODE
    assert "ExitApplication" in MASTER_SIDEBAR_FORM_CODE
    assert "Διαχείριση Μητρώων" in MASTER_SIDEBAR_FORM_CODE
    assert "Απουσίες / Ακυρώσεις" in MASTER_SIDEBAR_FORM_CODE
    assert "Ημερήσιο πρόγραμμα" in MASTER_SIDEBAR_FORM_CODE
    assert "Save & Exit" in MASTER_SIDEBAR_FORM_CODE


def test_master_sidebar_buttons_are_uniform_and_centered():
    assert '.Font.Name = "Segoe UI"' in MASTER_SIDEBAR_FORM_CODE
    assert ".Font.Size = 10" in MASTER_SIDEBAR_FORM_CODE
    assert ".TextAlign = 2" in MASTER_SIDEBAR_FORM_CODE
    assert ".WordWrap = False" in MASTER_SIDEBAR_FORM_CODE
    assert ".Width = 146" in MASTER_SIDEBAR_FORM_CODE
    assert ".Height = 36" in MASTER_SIDEBAR_FORM_CODE


def test_master_sidebar_is_shown_only_on_master_and_hidden_elsewhere():
    keep_start = APP_SHELL_MODULE_CODE.index("Public Sub KeepApplicationShell()")
    show_start = APP_SHELL_MODULE_CODE.index("Public Sub ShowMasterSidebar()")
    keep_block = APP_SHELL_MODULE_CODE[keep_start:show_start]

    assert 'If ActiveSheet.Name = "MASTER_SCHEDULE" Then' in keep_block
    assert "ShowMasterSidebar" in keep_block
    assert "HideMasterSidebar" in keep_block


def test_master_sidebar_hides_legacy_sheet_shapes_to_avoid_duplicates():
    assert "SetMasterToolbarShapesVisible False" in APP_SHELL_MODULE_CODE
    assert 'If Left$(shp.Name, 11) = "ekaToolbar_" Then' in APP_SHELL_MODULE_CODE


def test_master_sidebar_is_cleaned_up_when_app_shell_exits():
    exit_start = APP_SHELL_MODULE_CODE.index("Public Sub ExitApplicationShell()")
    restore_start = APP_SHELL_MODULE_CODE.index("Public Sub RestoreExcelInterface()")
    exit_block = APP_SHELL_MODULE_CODE[exit_start:restore_start]

    assert "HideMasterSidebar" in exit_block
    assert "Unload frmMasterSidebar" in exit_block


def test_master_sidebar_form_is_installed_with_four_buttons():
    source = inspect.getsource(_install_master_sidebar_form)
    assert "VBComponents.Add(3)" in source
    assert '"cmdRegistry"' in source
    assert '"cmdAbsences"' in source
    assert '"cmdDaily"' in source
    assert '"cmdExit"' in source
    assert "form.CodeModule.AddFromString(MASTER_SIDEBAR_FORM_CODE)" in source
