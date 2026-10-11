import inspect

from rehab_excel import patient_registration_scheduling_slice as scheduling_slice


def test_scheduling_ui_installs_combo_clear_handlers_after_program_form():
    source = inspect.getsource(scheduling_slice._install_scheduling_ui)
    assert "install_patient_program_ui" in source
    assert "install_patient_program_master_render(vbproject)" in source
    assert "install_combo_clear_handlers(vbproject)" in source
    assert source.index("install_patient_program_ui") < source.index(
        "install_patient_program_master_render(vbproject)"
    ) < source.index("install_combo_clear_handlers(vbproject)")


def test_scheduling_ui_reports_exact_installation_stage_on_com_failure():
    source = inspect.getsource(scheduling_slice._install_scheduling_ui)
    assert 'stage = "installing patient-program form/runtime"' in source
    assert 'stage = "installing/updating patient edit form"' in source
    assert 'stage = "installing MASTER patient-program renderer"' in source
    assert 'stage = "installing combo clear handlers"' in source
    assert 'stage = "saving preview workbook"' in source
    assert 'NumberFormat = "hh:mm"' not in source
    assert "failed during {stage}" in source
