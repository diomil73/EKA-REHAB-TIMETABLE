from types import SimpleNamespace

from rehab_excel.patient_program_conflict_layout_vba import (
    CONFLICT_FORM_RUNTIME_CODE,
    install_conflict_runtime_layout,
)


def test_conflict_form_runtime_layout_sizes_message_and_buttons():
    code = CONFLICT_FORM_RUNTIME_CODE
    assert "Private Sub ApplyRuntimeLayout()" in code
    assert "Me.Width = 620" in code
    assert "Me.Height = 390" in code
    assert "lblConflict.Height = 285" in code
    assert "cmdContinue.Top = 320" in code
    assert "cmdReturn.Top = 320" in code
    assert "ApplyRuntimeLayout" in code
    assert 'cmdContinue_Click' in code
    assert 'cmdReturn_Click' in code


def test_conflict_runtime_layout_replaces_existing_form_code():
    class FakeCodeModule:
        def __init__(self):
            self.text = "old code"

        @property
        def CountOfLines(self):
            return 1

        def DeleteLines(self, start, count):
            assert start == 1
            assert count == 1
            self.text = ""

        def AddFromString(self, text):
            self.text = text

    module = FakeCodeModule()
    form = SimpleNamespace(CodeModule=module)

    class FakeComponents:
        def __call__(self, name):
            assert name == "frmPatientProgramConflict"
            return form

    vbproject = SimpleNamespace(VBComponents=FakeComponents())
    install_conflict_runtime_layout(vbproject)

    assert module.text == CONFLICT_FORM_RUNTIME_CODE
