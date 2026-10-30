import inspect

from tools import build_registration_menu_preview as module


def test_outpatient_menu_diagnostic_handler_exposes_native_error_details():
    handler = module.OUTPATIENT_DIAGNOSTIC_HANDLER

    assert "frmOutpatientSchedule.Show" in handler
    assert "Err.Number" in handler
    assert "Err.Description" in handler
    assert "δεν άνοιξε" in handler


def test_canonical_registration_build_patches_and_verifies_outpatient_handler():
    source = inspect.getsource(module)

    assert "_replace_outpatient_menu_handler(vbproject)" in source
    assert 'ProcStartLine(proc_name, 0)' in source
    assert 'ProcCountLines(proc_name, 0)' in source
    assert 'code.DeleteLines(start_line, line_count)' in source
    assert 'code.InsertLines(start_line, OUTPATIENT_DIAGNOSTIC_HANDLER)' in source
    assert '"Err.Description" not in installed' in source
    assert '"Err.Number" not in installed' in source


def test_canonical_registration_build_verifies_global_replacement_action():
    source = inspect.getsource(module)

    assert "GLOBAL_REPLACEMENT_MODULE_NAME" in source
    assert 'stage = "verifying global replacement action"' in source
    assert "vbproject.VBComponents(GLOBAL_REPLACEMENT_MODULE_NAME)" in source
