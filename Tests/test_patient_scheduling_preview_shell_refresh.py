from pathlib import Path


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "Python"
    / "tools"
    / "build_patient_scheduling_slice_preview.py"
).read_text(encoding="utf-8")


def test_preview_builder_refreshes_current_application_shell_after_main_build():
    assert "def _refresh_application_shell" in SCRIPT
    assert "install_application_shell(vbproject, workbook=workbook)" in SCRIPT
    assert "_refresh_application_shell(args.output)" in SCRIPT


def test_preview_builder_verifies_sidebar_macro_and_form_before_saving():
    assert "APP_SHELL_MODULE_NAME" in SCRIPT
    assert "MASTER_SIDEBAR_FORM_NAME" in SCRIPT
    assert '"Public Sub ShowMasterSidebar()"' in SCRIPT
    assert "MASTER sidebar form missing after shell refresh" in SCRIPT
