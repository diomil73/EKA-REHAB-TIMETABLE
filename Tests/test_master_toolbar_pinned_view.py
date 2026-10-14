from rehab_excel.application_shell_vba import APP_SHELL_MODULE_CODE


def test_master_toolbar_is_pinned_above_scrollable_program():
    assert "Public Sub PinMasterToolbar()" in APP_SHELL_MODULE_CODE
    assert ".FreezePanes = False" in APP_SHELL_MODULE_CODE
    assert ".SplitColumn = 0" in APP_SHELL_MODULE_CODE
    assert ".SplitRow = 3" in APP_SHELL_MODULE_CODE
    assert ".FreezePanes = True" in APP_SHELL_MODULE_CODE


def test_master_toolbar_pinning_is_reapplied_on_shell_entry_and_return():
    enter_start = APP_SHELL_MODULE_CODE.index("Public Sub EnterApplicationShell()")
    keep_start = APP_SHELL_MODULE_CODE.index("Public Sub KeepApplicationShell()")
    pin_start = APP_SHELL_MODULE_CODE.index("Public Sub PinMasterToolbar()")

    enter_block = APP_SHELL_MODULE_CODE[enter_start:keep_start]
    keep_block = APP_SHELL_MODULE_CODE[keep_start:pin_start]

    assert "ApplyMasterView" in enter_block
    assert "PinMasterToolbar" in enter_block
    assert "ApplyMasterView" in keep_block
    assert "PinMasterToolbar" in keep_block


def test_master_toolbar_pinning_only_applies_to_master_sheet():
    pin_start = APP_SHELL_MODULE_CODE.index("Public Sub PinMasterToolbar()")
    chrome_start = APP_SHELL_MODULE_CODE.index("Private Sub ApplyApplicationChrome()")
    pin_block = APP_SHELL_MODULE_CODE[pin_start:chrome_start]

    assert 'If ActiveSheet.Name <> "MASTER_SCHEDULE" Then Exit Sub' in pin_block
