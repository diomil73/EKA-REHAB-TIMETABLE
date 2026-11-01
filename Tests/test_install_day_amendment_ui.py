import inspect

from tools import install_day_amendment_ui as module


def test_ui_installer_uses_dedicated_amendment_action_installer():
    source = inspect.getsource(module.install_day_amendment_ui)

    assert "install_day_amendment_action(vbproject, workbook)" in source
    assert "DAY_AMENDMENT_MODULE_NAME" in source
    assert "DAY_AMENDMENT_BUTTON_NAME" in source
    assert 'workbook.Worksheets("MASTER_SCHEDULE").Shapes(DAY_AMENDMENT_BUTTON_NAME)' in source
    assert 'str(button.OnAction) != "StartClosedDayAmendment"' in source
    assert "workbook.Save()" in source


def test_ui_installer_is_fail_closed_for_non_xlsm_or_missing_vba():
    source = inspect.getsource(module.install_day_amendment_ui)

    assert 'path.suffix.casefold() != ".xlsm"' in source
    assert "if not bool(workbook.HasVBProject)" in source
    assert "Day-amendment VBA module was not confirmed" in source
    assert "Day-amendment MASTER button was not confirmed" in source
