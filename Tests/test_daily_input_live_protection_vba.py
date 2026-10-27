import inspect

from rehab_excel import daily_input_live_protection_vba as protection
from tools import repair_daily_input_live_protection as repair_tool


def test_live_protection_module_reapplies_session_scoped_ui_only_mode():
    assert 'ws.Protect UserInterfaceOnly:=True' in protection.MODULE_CODE
    assert 'ws.Range("A7:F16").Locked = False' in protection.MODULE_CODE
    assert 'ws.Range("A21:E40").Locked = False' in protection.MODULE_CODE


def test_live_protection_locks_only_active_daily_input_rows_from_ledger_metadata():
    source = protection.MODULE_CODE
    assert 'UCase$(Trim$(CStr(ledger.Cells(row, 12).Value))) = STATUS_ACTIVE' in source
    assert 'Left$(meta, Len(META_PREFIX)) = META_PREFIX' in source
    assert 'InStrRev(meta, ":") + 1' in source
    assert 'sourceRow >= 7 And sourceRow <= 16' in source
    assert 'sourceRow >= 21 And sourceRow <= 40' in source


def test_installer_hooks_ledger_sync_and_manual_projection_fail_closed():
    source = inspect.getsource(protection.install_daily_input_live_protection)
    assert 'DAILY_INPUT ledger live-protection hook' in source
    assert 'replacement operator live-protection hook' in source
    assert 'RefreshDailyInputProtectionFromLedger' in protection._LEDGER_REPLACEMENT
    assert 'operator projection protection' in protection._OPERATOR_REPLACEMENT


def test_repair_tool_uses_rollback_and_verifies_all_hooks():
    source = inspect.getsource(repair_tool.repair_workbook)
    assert 'shutil.copy2(path, rollback)' in source
    assert 'install_daily_input_live_protection(workbook.VBProject)' in source
    assert '_verify(workbook.VBProject)' in source
    assert 'shutil.copy2(rollback, path)' in source
