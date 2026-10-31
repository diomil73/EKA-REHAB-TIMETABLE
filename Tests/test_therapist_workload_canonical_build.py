import inspect

from tools import build_registration_menu_preview


def test_canonical_registration_build_installs_and_verifies_workload_strip():
    install_source = inspect.getsource(
        build_registration_menu_preview._install_close_day_action_in_workbook
    )
    main_source = inspect.getsource(build_registration_menu_preview.main)
    clean_source = inspect.getsource(
        build_registration_menu_preview._prepare_clean_vba_copy
    )

    assert "install_therapist_workload_strip(vbproject, workbook)" in install_source
    assert "WORKLOAD_MODULE_NAME" in install_source
    assert '("verifying therapist workload strip", WORKLOAD_MODULE_NAME, "Therapist workload strip")' in install_source
    assert 'f"{error_label} module was not confirmed in the preview VBA project"' in install_source
    assert "_install_close_day_action_in_workbook(output)" in main_source
    assert "Therapist workload strip present:" in main_source
    assert '("removing old therapist workload strip module", WORKLOAD_MODULE_NAME)' in clean_source
