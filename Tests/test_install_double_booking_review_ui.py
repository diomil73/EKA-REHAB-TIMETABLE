from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "install_double_booking_review_ui.py"
PROJECTION_PATH = (
    REPO_ROOT / "Python" / "rehab_excel" / "double_booking_daily_projection_vba.py"
)


def test_ui_installer_uses_fresh_copy_and_application_contract():
    source = TOOL_PATH.read_text(encoding="utf-8")
    assert "shutil.copy2(source, output)" in source
    assert "assert_application_contract(source)" in source
    assert "assert_application_contract(output)" in source
    assert "file_sha256(source)" in source
    assert "output.unlink()" in source


def test_ui_installer_confirms_form_module_and_projection_module():
    source = TOOL_PATH.read_text(encoding="utf-8")
    assert "install_double_booking_review(vbproject)" in source
    assert "install_double_booking_daily_projection(vbproject)" in source
    assert "DOUBLE_BOOKING_MODULE_NAME" in source
    assert "DOUBLE_BOOKING_FORM_NAME" in source
    assert "DOUBLE_BOOKING_PROJECTION_MODULE_NAME" in source
    assert "DOUBLE BOOKING REVIEW UI PREVIEW OK" in source
    assert "Macro: ShowDoubleBookingReview" in source


def test_form_wiring_projects_selected_patient_into_effective_daily_slot():
    source = TOOL_PATH.read_text(encoding="utf-8")
    assert "modDoubleBookingDailyProjection.ProjectDoubleBookingToTherapistDaily" in source
    assert "lstOptions.List(idx, 1)" in source  # patient display name
    assert "lstOptions.List(idx, 4)" in source  # therapist display name
    assert "lstOptions.List(idx, 5)" in source  # replacement time
    assert "lstOptions.List(idx, 0)" in source  # target session for rollback


def test_projection_appends_instead_of_overwriting_and_rolls_back_on_failure():
    source = PROJECTION_PATH.read_text(encoding="utf-8")
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY")' in source
    assert 'targetCell.Value = existingText & vbLf & patientName' in source
    assert "TextHasExactLine" in source
    assert "ResolveTherapistDailyCell" in source
    assert "headerRows = Array(1, 11)" in source
    assert "firstRows = Array(2, 12)" in source
    assert "lastRows = Array(8, 18)" in source
    assert "RollbackValidatedFact targetSessionId" in source
    assert 'Trim$(CStr(ws.Cells(row, 8).Value)) = "CONFIRMED DOUBLE BOOKING EXCEPTION"' in source
    assert "ThisWorkbook.Save" in source
