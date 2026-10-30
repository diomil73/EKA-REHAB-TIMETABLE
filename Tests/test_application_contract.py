from __future__ import annotations

from pathlib import Path
import inspect

from rehab_excel.application_contract import (
    REQUIRED_APP_COMPONENTS,
    assert_application_contract,
)
from tools import preview_daily_input_sheet


def test_application_contract_covers_existing_app_ui_stack():
    required = set(REQUIRED_APP_COMPONENTS)

    expected = {
        "modApplicationShell",
        "modMasterToolbar",
        "modRegistrationMenu",
        "frmRegistrationMenu",
        "frmNewPatient",
        "frmNewTherapist",
        "frmNewStudent",
        "modPatientRegistrationBridge",
        "frmEditPatient",
        "modDailyInputAction",
        "modTherapistWorkloadStrip",
        "modOutpatientSchedule",
        "frmOutpatientSchedule",
        "modOutpatientInSession",
        "modSchedulingConflict",
        "frmSchedulingConflict",
        "modOutpatientScheduleEdit",
        "modOutpatientScheduleManager",
        "frmOutpatientScheduleManager",
    }

    assert expected <= required


def test_daily_input_preview_defaults_to_known_good_app_source():
    source = inspect.getsource(preview_daily_input_sheet)

    assert '"OUTPATIENT_RECURRING_SMOKE.xlsm"' in source
    assert 'default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm"' not in source
    assert "assert_application_contract(source)" in source
    assert "assert_application_contract(preview)" in source


def test_application_contract_requires_windows_excel_component_inspection():
    source = inspect.getsource(assert_application_contract)

    assert 'sys.platform != "win32"' in source
    assert 'DispatchEx("Excel.Application")' in source
    assert "workbook.VBProject" in source
    assert "REQUIRED_APP_COMPONENTS" in source
