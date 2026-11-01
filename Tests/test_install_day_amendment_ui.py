import inspect

from tools import install_day_amendment_ui as module


def test_ui_installer_reapplies_clean_master_before_day_save_action():
    source = inspect.getsource(module.install_day_amendment_ui)

    assert "install_master_toolbar(vbproject, workbook)" in source
    assert "install_day_amendment_action(vbproject, workbook)" in source
    assert source.index("install_master_toolbar") < source.index("install_day_amendment_action")
    assert "DAY_AMENDMENT_MODULE_NAME" in source
    assert "DAY_AMENDMENT_BUTTON_NAME" in source
    assert 'workbook.Worksheets("DAILY_INPUT").Shapes(DAY_AMENDMENT_BUTTON_NAME)' in source
    assert 'str(button.OnAction) != "StartClosedDayAmendment"' in source
    assert "master_shapes = _master_toolbar_shapes(workbook)" in source
    assert "if master_shapes:" in source
    assert "workbook.Save()" in source


def test_ui_installer_is_fail_closed_for_non_xlsm_missing_vba_or_wrong_placement():
    source = inspect.getsource(module.install_day_amendment_ui)

    assert 'path.suffix.casefold() != ".xlsm"' in source
    assert "if not bool(workbook.HasVBProject)" in source
    assert "Day-amendment VBA module was not confirmed" in source
    assert "MASTER_SCHEDULE still contains operational toolbar shapes" in source
    assert "Day-save DAILY_INPUT button was not confirmed" in source


def test_master_shape_audit_only_targets_operational_toolbar_prefix():
    source = inspect.getsource(module._master_toolbar_shapes)

    assert 'workbook.Worksheets("MASTER_SCHEDULE")' in source
    assert "startswith(MASTER_TOOLBAR_PREFIX)" in source
