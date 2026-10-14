from rehab_excel.application_shell_vba import APP_SHELL_MODULE_CODE


def test_master_sidebar_is_scroll_independent_command_bar():
    assert 'MASTER_SIDEBAR_NAME As String = "EKA_MASTER_SIDEBAR"' in APP_SHELL_MODULE_CODE
    assert "Public Sub ShowMasterSidebar()" in APP_SHELL_MODULE_CODE
    assert "Private Sub BuildMasterSidebar()" in APP_SHELL_MODULE_CODE
    assert "Application.CommandBars.Add" in APP_SHELL_MODULE_CODE
    assert "Position:=0" in APP_SHELL_MODULE_CODE


def test_master_sidebar_contains_all_four_master_actions():
    assert '"Διαχείριση Μητρώων", "ToolbarRegistryManagement"' in APP_SHELL_MODULE_CODE
    assert '"Απουσίες / Ακυρώσεις", "ToolbarAbsences"' in APP_SHELL_MODULE_CODE
    assert '"Ημερήσιο πρόγραμμα", "ToolbarTherapistDaily"' in APP_SHELL_MODULE_CODE
    assert '"Save & Exit", "ExitApplication"' in APP_SHELL_MODULE_CODE


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
    assert "Application.CommandBars(MASTER_SIDEBAR_NAME).Delete" in exit_block
