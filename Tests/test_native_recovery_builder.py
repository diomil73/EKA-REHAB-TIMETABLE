import inspect

from tools import build_unified_preview_native_recovery


def test_native_recovery_builder_applies_patch_after_unified_build():
    source = inspect.getsource(
        build_unified_preview_native_recovery.build_native_recovery_preview
    )
    assert "build_unified_preview(" in source
    assert "patch_workbook(built)" in source
    assert source.index("build_unified_preview(") < source.index("patch_workbook(built)")
    assert "assert_application_contract(built)" in source
    assert source.index("patch_workbook(built)") < source.index(
        "assert_application_contract(built)"
    )
    assert "output.unlink(missing_ok=True)" in source
    assert "Original source workbook changed" in source


def test_native_recovery_builder_reports_patch_verification():
    source = inspect.getsource(build_unified_preview_native_recovery.main)
    assert '"Native M recovery patch: verified"' in source
    assert '"NATIVE RECOVERY UNIFIED PREVIEW OK"' in source
