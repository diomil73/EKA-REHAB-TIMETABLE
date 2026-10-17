import inspect

from rehab_excel.application_shell_vba import (
    APP_SHELL_MODULE_CODE,
    RETIRED_MASTER_SIDEBAR_FORM_NAME,
    install_application_shell,
)


def test_application_shell_uses_only_targeted_master_freeze_pinning():
    assert "ShowMasterSidebar" not in APP_SHELL_MODULE_CODE
    assert "HideMasterSidebar" not in APP_SHELL_MODULE_CODE
    assert "vbModeless" not in APP_SHELL_MODULE_CODE
    assert "Application.CommandBars.Add" not in APP_SHELL_MODULE_CODE
    assert "Private Sub PinMasterToolbarAndHeaders()" in APP_SHELL_MODULE_CODE
    assert ".SplitColumn = 0" in APP_SHELL_MODULE_CODE
    assert ".SplitRow = 4" in APP_SHELL_MODULE_CODE
    assert ".FreezePanes = True" in APP_SHELL_MODULE_CODE


def test_application_shell_keeps_master_view_and_reapplies_pin():
    keep_start = APP_SHELL_MODULE_CODE.index("Public Sub KeepApplicationShell()")
    keep_end = APP_SHELL_MODULE_CODE.index("End Sub", keep_start)
    keep_block = APP_SHELL_MODULE_CODE[keep_start:keep_end]

    assert 'If ActiveSheet.Name = "MASTER_SCHEDULE" Then' in keep_block
    assert "ApplyMasterView" in keep_block
    assert "PinMasterToolbarAndHeaders" in keep_block
    assert "ShowMasterSidebar" not in keep_block
    assert "HideMasterSidebar" not in keep_block


def test_retired_master_sidebar_form_is_explicitly_removed_on_install():
    assert RETIRED_MASTER_SIDEBAR_FORM_NAME == "frmMasterSidebar"
    source = inspect.getsource(install_application_shell)
    assert "_remove_component_if_present(vbproject, RETIRED_MASTER_SIDEBAR_FORM_NAME)" in source


def test_application_shell_still_restores_excel_interface_on_exit():
    assert "Public Sub ExitApplicationShell()" in APP_SHELL_MODULE_CODE
    assert "Public Sub RestoreExcelInterface()" in APP_SHELL_MODULE_CODE
    assert "SetRibbonVisible previousRibbonVisible" in APP_SHELL_MODULE_CODE
