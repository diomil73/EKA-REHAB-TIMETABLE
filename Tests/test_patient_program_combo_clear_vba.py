from types import SimpleNamespace

from rehab_excel.patient_program_combo_clear_vba import (
    COMBO_CLEAR_HANDLER_CODE,
    PATIENT_PROGRAM_FORM_NAME,
    install_combo_clear_handlers,
)


def test_combo_clear_handlers_cover_days_time_and_provider():
    code = COMBO_CLEAR_HANDLER_CODE
    assert "cboDayPattern_KeyDown" in code
    assert "cboTime_KeyDown" in code
    assert "cboProvider_KeyDown" in code
    assert "vbKeyDelete" in code
    assert "vbKeyBack" in code
    assert "combo.ListIndex = -1" in code
    assert 'combo.Value = ""' in code
    assert "SaveCurrentDraft" in code
    assert "UpdateDraftStatus" in code
    assert "KeyCode = 0" in code


def test_combo_clear_installer_appends_to_patient_program_form():
    appended = []

    class FakeCodeModule:
        def AddFromString(self, code):
            appended.append(code)

    form = SimpleNamespace(CodeModule=FakeCodeModule())

    class FakeComponents:
        def __call__(self, name):
            assert name == PATIENT_PROGRAM_FORM_NAME
            return form

    vbproject = SimpleNamespace(VBComponents=FakeComponents())
    install_combo_clear_handlers(vbproject)

    assert appended == [COMBO_CLEAR_HANDLER_CODE]
