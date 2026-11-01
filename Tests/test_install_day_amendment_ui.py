import inspect

from tools import install_day_amendment_ui as module


def test_ui_installer_only_installs_day_save_action():
    source = inspect.getsource(module.install_day_amendment_ui)

    assert "install_day_amendment_action(vbproject, workbook)" in source
    assert "install_master_toolbar" not in source
    assert "MASTER_TOOLBAR_PREFIX" not in source
    assert "DAY_AMENDMENT_MODULE_NAME" in source
    assert "DAY_AMENDMENT_BUTTON_NAME" in source
    assert 'workbook.Worksheets("DAILY_INPUT").Shapes(DAY_AMENDMENT_BUTTON_NAME)' in source
    assert 'str(button.OnAction) != "StartClosedDayAmendment"' in source
    assert "workbook.Save()" in source


def test_ui_installer_is_fail_closed_for_non_xlsm_missing_vba_or_wrong_day_save_placement():
    source = inspect.getsource(module.install_day_amendment_ui)

    assert 'path.suffix.casefold() != ".xlsm"' in source
    assert "if not bool(workbook.HasVBProject)" in source
    assert "Day-amendment VBA module was not confirmed" in source
    assert "Day-save DAILY_INPUT button was not confirmed" in source


def test_cli_reports_daily_input_as_button_sheet_without_master_claims():
    source = inspect.getsource(module.main)

    assert 'print("Button sheet: DAILY_INPUT")' in source
    assert "MASTER clean" not in source
