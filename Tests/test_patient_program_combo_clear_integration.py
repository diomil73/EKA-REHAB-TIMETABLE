import inspect

from rehab_excel import patient_registration_scheduling_slice as scheduling_slice


def test_scheduling_ui_installs_combo_clear_handlers_after_program_form():
    source = inspect.getsource(scheduling_slice._install_scheduling_ui)
    assert "install_patient_program_ui" in source
    assert "install_combo_clear_handlers(workbook.VBProject)" in source
    assert source.index("install_patient_program_ui") < source.index(
        "install_combo_clear_handlers(workbook.VBProject)"
    )
