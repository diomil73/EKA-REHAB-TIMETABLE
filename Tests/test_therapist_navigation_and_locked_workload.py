import inspect

from rehab_excel.daily_input_change_navigation_vba import (
    MACRO_NAME as LEDGER_NAV_MACRO,
    MODULE_CODE as LEDGER_NAV_CODE,
    install_daily_input_change_navigation,
)
from rehab_excel.therapist_workload_strip_vba import (
    MODULE_CODE,
    install_therapist_workload_strip,
)


def test_supervisor_workload_opens_therapist_daily_and_return_to_master_exists():
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY").Activate' in MODULE_CODE
    assert "Public Sub GoToMasterFromTherapistDaily()" in MODULE_CODE
    assert '"Επιστροφή στο MASTER"' in MODULE_CODE
    assert 'shape.OnAction = "GoToMasterFromTherapistDaily"' in MODULE_CODE


def test_return_to_master_navigation_tracks_restore_panel():
    assert 'Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"' in MODULE_CODE
    assert 'Set ledger = ThisWorkbook.Worksheets(LEDGER_SHEET)' in MODULE_CODE
    assert 'ledger.Range("R1").Value' in MODULE_CODE
    assert 'navRow = restoreStartRow - 1' in MODULE_CODE
    assert 'daily.Range("A" & CStr(navRow))' in MODULE_CODE
    assert 'daily.Range("K1")' not in MODULE_CODE

    installer_source = inspect.getsource(install_therapist_workload_strip)
    assert "anchor = _navigation_anchor(daily, workbook)" in installer_source
    assert "float(anchor.Left) + 4" in installer_source
    assert "float(anchor.Top) + 2" in installer_source
    assert 'daily.Range("K1")' not in installer_source


def test_master_never_renders_or_exposes_therapist_workload():
    assert "Privacy rule: MASTER must never expose therapist workload." in MODULE_CODE
    assert "DeleteWorkloadShapes master" in MODULE_CODE
    assert 'shape.TextFrame2.TextRange.Text = "Φόρτος θεραπευτών"' not in MODULE_CODE

    installer_source = inspect.getsource(install_therapist_workload_strip)
    assert "_delete_workload_shapes(master)" in installer_source
    assert "master.Shapes.AddShape" not in installer_source


def test_workload_badges_render_only_on_therapist_daily_and_are_locked():
    assert "Set shape = daily.Shapes.AddShape" in MODULE_CODE
    assert 'leftPos = daily.Range("K2").Left + 4' in MODULE_CODE
    assert "shape.Locked = True" in MODULE_CODE


def test_supervisor_entry_point_syncs_reversible_daily_changes_without_master_binding():
    assert LEDGER_NAV_MACRO == "OpenTherapistDailyWithChanges"
    assert "ShowTherapistWorkload" in LEDGER_NAV_CODE
    assert "SyncDailyInputChangesToLedger" in LEDGER_NAV_CODE

    installer_source = inspect.getsource(install_daily_input_change_navigation)
    assert 'Worksheets("MASTER_SCHEDULE")' not in installer_source
    assert "Shapes(" not in installer_source
