from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "install_double_booking_review_ui.py"


def test_ui_installer_uses_fresh_copy_and_application_contract():
    source = TOOL_PATH.read_text(encoding="utf-8")
    assert "shutil.copy2(source, output)" in source
    assert "assert_application_contract(source)" in source
    assert "assert_application_contract(output)" in source
    assert "file_sha256(source)" in source
    assert "output.unlink()" in source


def test_ui_installer_confirms_form_and_module():
    source = TOOL_PATH.read_text(encoding="utf-8")
    assert "install_double_booking_review(vbproject)" in source
    assert "DOUBLE_BOOKING_MODULE_NAME" in source
    assert "DOUBLE_BOOKING_FORM_NAME" in source
    assert "DOUBLE BOOKING REVIEW UI PREVIEW OK" in source
    assert "Macro: ShowDoubleBookingReview" in source
