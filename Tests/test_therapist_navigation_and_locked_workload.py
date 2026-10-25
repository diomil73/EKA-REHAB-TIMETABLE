from rehab_excel.daily_input_change_navigation_vba import (
    MACRO_NAME as LEDGER_NAV_MACRO,
    MODULE_CODE as LEDGER_NAV_CODE,
)
from rehab_excel.therapist_workload_strip_vba import MODULE_CODE


def test_workload_button_opens_therapist_daily_and_return_to_master_exists():
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY").Activate' in MODULE_CODE
    assert "Public Sub GoToMasterFromTherapistDaily()" in MODULE_CODE
    assert '"Επιστροφή στο MASTER"' in MODULE_CODE
    assert 'shape.OnAction = "GoToMasterFromTherapistDaily"' in MODULE_CODE


def test_workload_shapes_are_locked_and_master_drawing_objects_protected():
    assert "shape.Locked = True" in MODULE_CODE
    assert "DrawingObjects:=True" in MODULE_CODE
    assert "UserInterfaceOnly:=True" in MODULE_CODE


def test_workload_entry_point_syncs_reversible_daily_changes():
    assert LEDGER_NAV_MACRO == "OpenTherapistDailyWithChanges"
    assert "ShowTherapistWorkload" in LEDGER_NAV_CODE
    assert "SyncDailyInputChangesToLedger" in LEDGER_NAV_CODE
