import inspect

from rehab_excel.daily_input_change_navigation_vba import (
    MODULE_CODE,
    STYLE_MODULE_CODE,
    STYLE_MODULE_NAME,
    install_daily_input_change_navigation,
)


def test_supervisor_navigation_repaints_after_ledger_sync():
    assert "SyncDailyInputChangesToLedger" in MODULE_CODE
    assert "StyleSupervisorPanel" in MODULE_CODE


def test_supervisor_panel_uses_application_shapes_not_grey_form_buttons():
    assert STYLE_MODULE_NAME == "modSupervisorPanelStyle"
    assert "Public Sub StyleSupervisorPanel()" in STYLE_MODULE_CODE
    assert "Shapes.AddShape" in STYLE_MODULE_CODE
    assert "AddFormControl" not in STYLE_MODULE_CODE
    assert '"Επιστροφή στο MASTER"' in STYLE_MODULE_CODE
    assert '"Επαναφορά επιλεγμένων"' in STYLE_MODULE_CODE
    assert '"Επαναφορά όλων"' in STYLE_MODULE_CODE


def test_supervisor_panel_has_mint_and_coral_visual_hierarchy():
    assert "RGB(239, 249, 246)" in STYLE_MODULE_CODE
    assert "RGB(204, 235, 226)" in STYLE_MODULE_CODE
    assert "RGB(250, 226, 229)" in STYLE_MODULE_CODE
    assert "RGB(248, 216, 219)" in STYLE_MODULE_CODE
    assert "RGB(241, 196, 199)" in STYLE_MODULE_CODE
    assert "RowHeight = 30" in STYLE_MODULE_CODE
    assert "RowHeight = 38" in STYLE_MODULE_CODE


def test_installer_injects_style_after_all_ledger_refresh_paths():
    source = inspect.getsource(install_daily_input_change_navigation)
    assert "STYLE_MODULE_NAME" in source
    assert "STYLE_MODULE_CODE" in source
    assert "_inject_style_after_refresh(vbproject)" in source
    assert "_inject_style_after_daily_input_sync(vbproject)" in source
