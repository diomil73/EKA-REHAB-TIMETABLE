import inspect

from rehab_excel.master_toolbar_vba import MASTER_TOOLBAR_MODULE_CODE
from rehab_excel.patient_history_report_vba import PATIENT_REPORT_VBA
from rehab_excel.therapist_history_report_vba import (
    THERAPIST_HISTORY_MODULE_CODE,
    install_therapist_history_report,
)


def test_existing_previous_screen_macro_returns_to_therapist_daily():
    assert "Public Sub ToolbarTherapistDaily()" in MASTER_TOOLBAR_MODULE_CODE
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY").Activate' in MASTER_TOOLBAR_MODULE_CODE


def test_patient_report_back_button_returns_directly_to_previous_screen():
    src = PATIENT_REPORT_VBA
    assert 'EnsurePatientButton ws, "Back", "← Προηγούμενη", "ToolbarTherapistDaily", 116#' in src
    assert 'EnsurePatientButton ws, "Back", "← MASTER"' not in src
    assert '"GoToMaster", 86#' not in src


def test_therapist_report_back_button_same_destination_and_button_width():
    src = THERAPIST_HISTORY_MODULE_CODE
    assert 'EnsureOneReportButton ws, "ekaHistory_Back", "← Προηγούμενη", "ToolbarTherapistDaily", 116#' in src
    assert 'Array("ekaHistory_Back", 116#)' in src
    installer = inspect.getsource(install_therapist_history_report)
    assert '("Back", "← Προηγούμενη", "ToolbarTherapistDaily")' in installer
    assert 'width = 90.0 if suffix != "Back" else 116.0' in installer
    assert 'EnsureOneReportButton ws, "ekaHistory_Back", "← MASTER"' not in src
