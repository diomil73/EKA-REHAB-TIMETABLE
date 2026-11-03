import inspect

from tools import build_patient_history_h5_smoke as module


def test_h5_smoke_builder_installs_visible_close_and_save_controls():
    source = inspect.getsource(module._install_and_verify_h5_controls)

    assert 'ws.Range("A4")' in source
    assert 'ws.Range("D4")' in source
    assert '"Κλείσιμο ημέρας"' in source
    assert '"Αποθήκευση ημέρας"' in source
    assert "CLOSE_DAY_MACRO_NAME" in source
    assert "DAY_AMENDMENT_MACRO_NAME" in source
    assert "verified_close.OnAction" in source
    assert "verified_save.OnAction" in source
    assert "verified_close.Visible" in source
    assert "verified_save.Visible" in source


def test_h5_smoke_builder_still_requires_zero_preexisting_history():
    source = inspect.getsource(module.build_patient_history_h5_smoke)

    assert source.count("_assert_clean_history(output)") >= 2
    assert "install_day_amendment_ui(output)" in source
    assert "_install_and_verify_h5_controls(output)" in source


def test_h5_smoke_success_output_reports_verified_buttons():
    source = inspect.getsource(module.main)

    assert "Close-day button: visible + verified" in source
    assert "Day-save button: visible + verified" in source
