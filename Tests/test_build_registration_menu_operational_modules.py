import inspect

from tools import build_registration_menu_preview as module


def test_preview_builder_installs_operational_modules_after_shell_build():
    source = inspect.getsource(module._install_operational_modules)

    assert "install_global_replacement_action(vbproject)" in source
    assert "install_double_booking_review(vbproject)" in source
    assert "install_therapist_workload_strip(vbproject, workbook)" in source
    assert "GLOBAL_REPLACEMENT_MODULE_NAME" in source
    assert "DOUBLE_BOOKING_MODULE_NAME" in source
    assert "DOUBLE_BOOKING_FORM_NAME" in source
    assert "WORKLOAD_MODULE_NAME" in source
    assert "missing = [name for name in required" in source


def test_preview_builder_runs_operational_install_before_source_hash_gate():
    source = inspect.getsource(module.main)
    build_pos = source.index("create_registration_menu_preview(")
    operational_pos = source.index("_install_operational_modules(output)")
    hash_pos = source.index("if _sha256(source) != source_before")

    assert build_pos < operational_pos < hash_pos


def test_preclean_removes_operational_modules_before_reinstall():
    source = inspect.getsource(module._prepare_clean_vba_copy)

    assert "GLOBAL_REPLACEMENT_MODULE_NAME" in source
    assert "DOUBLE_BOOKING_MODULE_NAME" in source
    assert "DOUBLE_BOOKING_FORM_NAME" in source
    assert "WORKLOAD_MODULE_NAME" in source
