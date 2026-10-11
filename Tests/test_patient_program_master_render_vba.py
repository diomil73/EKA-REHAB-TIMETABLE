from rehab_excel.patient_program_master_render_vba import (
    MASTER_RENDER_MODULE_CODE,
    MASTER_RENDER_MODULE_NAME,
    _patch_code_module,
    _patch_patient_edit_refresh,
)


def test_master_renderer_builds_time_days_and_optional_provider_lines():
    code = MASTER_RENDER_MODULE_CODE
    assert MASTER_RENDER_MODULE_NAME == "PatientProgramMasterRender"
    assert 'timeCol = FindHeaderColumn(planner, treatment & "_Ώρα")' in code
    assert 'daysCol = FindHeaderColumn(planner, treatment & "_Ημέρες")' in code
    assert 'ProviderHeaderName = treatment & "_Θεραπευτής"' in code
    assert 'ProviderHeaderName = "Ψυχ_Ψυχολόγος"' in code
    assert 'TimeText = Format$(rawValue, "hh:mm")' in code
    assert "result = result & vbLf" in code


def test_master_renderer_uses_general_format_for_composite_text_not_hhmm_number_format():
    code = MASTER_RENDER_MODULE_CODE
    assert '.NumberFormat = "General"' in code
    assert '.Value = BuildTreatmentDisplay(planner, plannerRow, treatment)' in code
    assert '.WrapText = True' in code


def test_master_renderer_maps_all_patient_program_treatments_to_master_columns():
    code = MASTER_RENDER_MODULE_CODE
    expected = {
        4: "ΦΘ",
        5: "Ρομποτικό",
        6: "Πισίνα",
        7: "Ανακλινόμενο",
        8: "Εργο",
        9: "Λογο",
        10: "ΕΦΑ",
    }
    for column, treatment in expected.items():
        assert f'WriteTreatmentCell master.Cells(masterRow, {column}), planner, plannerRow, "{treatment}"' in code
    assert '"ΨΥΧΟΛΟΓΟΙ"' in code
    assert 'WriteTreatmentCell master.Cells(masterRow, 11), planner, plannerRow, "Ψυχ"' in code


def _fake_code_module(text):
    class FakeCodeModule:
        def __init__(self, source):
            self.text = source
            self.add_count = 0

        @property
        def CountOfLines(self):
            return len(self.text.splitlines())

        def Lines(self, start, count):
            return self.text

        def DeleteLines(self, start, count):
            self.text = ""

        def AddFromString(self, source):
            self.text = source
            self.add_count += 1

    return FakeCodeModule(text)


def test_master_render_patcher_matches_vbe_crlf_code_with_lf_markers():
    module = _fake_code_module(
        "before\r\n"
        "    ThisWorkbook.Save\r\n"
        "    Application.EnableEvents = previousEvents\r\n"
        "after\r\n"
    )

    _patch_code_module(
        module,
        "    ThisWorkbook.Save\n    Application.EnableEvents = previousEvents",
        "    RefreshMasterPatientProgram mPatientId\n"
        "    ThisWorkbook.Save\n"
        "    Application.EnableEvents = previousEvents",
        description="patient-program save refresh",
    )

    assert "RefreshMasterPatientProgram mPatientId" in module.text
    assert "ThisWorkbook.Save" in module.text


def test_master_render_patcher_is_idempotent_after_refresh_is_present():
    module = _fake_code_module(
        "    RefreshMasterPatientProgram mPatientId\r\n"
        "    ThisWorkbook.Save\r\n"
        "    Application.EnableEvents = previousEvents\r\n"
    )

    _patch_code_module(
        module,
        "    ThisWorkbook.Save\n    Application.EnableEvents = previousEvents",
        "    RefreshMasterPatientProgram mPatientId\n"
        "    ThisWorkbook.Save\n"
        "    Application.EnableEvents = previousEvents",
        description="patient-program save refresh",
    )

    assert module.add_count == 0


def test_patient_edit_refresh_patch_is_scoped_to_cmdsave_and_tolerates_vbe_spacing():
    module = _fake_code_module(
        "Private Function ValidateForm() As Boolean\r\n"
        "    If Len(Trim$(cboStatus.Value)) = 0 Then Exit Function\r\n"
        "End Function\r\n"
        "\r\n"
        "Private Sub cmdSave_Click()\r\n"
        "    Dim patientId As String\r\n"
        "    patientId = Trim$(txtPatientID.Text)\r\n"
        "    UpdatePatientInWorkbook _\r\n"
        "        patientId, _\r\n"
        "        Trim$(cboStatus.Value)\r\n"
        "    MsgBox \"Οι αλλαγές αποθηκεύτηκαν επιτυχώς.\"\r\n"
        "End Sub\r\n"
    )

    _patch_patient_edit_refresh(module)

    save_block = module.text.split("Private Sub cmdSave_Click()", 1)[1]
    assert "RefreshMasterPatientProgram patientId" in save_block
    assert save_block.index("RefreshMasterPatientProgram patientId") < save_block.index(
        'MsgBox "Οι αλλαγές αποθηκεύτηκαν επιτυχώς."'
    )


def test_patient_edit_refresh_patch_is_idempotent():
    module = _fake_code_module(
        "Private Sub cmdSave_Click()\r\n"
        "    Dim patientId As String\r\n"
        "    RefreshMasterPatientProgram patientId\r\n"
        "    MsgBox \"Οι αλλαγές αποθηκεύτηκαν επιτυχώς.\"\r\n"
        "End Sub\r\n"
    )

    _patch_patient_edit_refresh(module)

    assert module.add_count == 0
