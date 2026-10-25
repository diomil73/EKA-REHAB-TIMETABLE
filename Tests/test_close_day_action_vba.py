import inspect

from rehab_excel.close_day_action_vba import (
    CLOSE_DAY_MACRO_NAME,
    CLOSE_DAY_MODULE_CODE,
    CLOSE_DAY_MODULE_NAME,
)
from rehab_excel.registration_menu_vba import Win32ComRegistrationMenuInstaller


def test_close_day_action_has_explicit_confirmation_and_authoritative_commit():
    source = CLOSE_DAY_MODULE_CODE

    assert CLOSE_DAY_MODULE_NAME == "modCloseDayAction"
    assert CLOSE_DAY_MACRO_NAME == "CloseOperationalDay"
    assert "vbYesNo + vbDefaultButton2" in source
    assert "ThisWorkbook.Save" in source
    assert "close_day.py" in source
    assert "authoritative_commit_worker_cli.py" in source
    assert '"post_reopen_target"' in source
    assert '"MASTER_SCHEDULE"' in source
    assert "ScheduleDeferredCloseDayHandoff" in source
    assert "Application.OnTime" in source
    assert "Public Sub FinishCloseDayHandoff()" in source
    assert "ThisWorkbook.Close SaveChanges:=False" in source
    assert "ThisWorkbook.Close SaveChanges:=True" not in source


def test_close_day_action_is_fail_closed_on_backend_error():
    source = CLOSE_DAY_MODULE_CODE

    assert 'If exec.ExitCode <> 0' in source
    assert 'JsonStringValue(responseText, "error")' in source
    assert 'Η ημέρα δεν έκλεισε' in source
    assert 'StartCloseDayCommit' in source


def test_close_day_handoff_exits_original_macro_before_closing_workbook():
    source = CLOSE_DAY_MODULE_CODE

    schedule_pos = source.index("ScheduleDeferredCloseDayHandoff")
    finish_pos = source.index("Public Sub FinishCloseDayHandoff()")
    close_pos = source.index("ThisWorkbook.Close SaveChanges:=False")
    assert schedule_pos < finish_pos < close_pos
    assert "Application.EnableEvents = False" in source


def test_registration_menu_installer_installs_and_verifies_close_day_action():
    source = inspect.getsource(Win32ComRegistrationMenuInstaller.install)

    assert "install_close_day_action(vbproject)" in source
    assert "CLOSE_DAY_MODULE_NAME" in source
    assert "Close-day action module was not confirmed" in source
