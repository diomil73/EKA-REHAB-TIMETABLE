from rehab_excel.patient_program_form_vba import (
    PATIENT_PROGRAM_FORM_CODE,
    PATIENT_PROGRAM_MODULE_CODE,
)
from rehab_excel.patient_program_master_render_vba import (
    CONFLICT_FORM_CODE,
    MASTER_RENDER_MODULE_CODE,
    MASTER_RENDER_MODULE_NAME,
    _ensure_conflict_control,
    _patch_code_module,
    _patch_patient_edit_refresh,
    _patch_program_conflict_validation,
    _patch_runtime_validation,
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


def test_master_renderer_guards_empty_patient_cells_before_split():
    code = MASTER_RENDER_MODULE_CODE
    start = code.index("Private Function FirstLine")
    end = code.index("Private Function FindMasterPatientRow", start)
    first_line = code[start:end]
    assert 'normalized = Trim$(Replace(valueText, vbCrLf, vbLf))' in first_line
    assert 'If Len(normalized) = 0 Then' in first_line
    assert 'FirstLine = ""' in first_line
    assert 'Exit Function' in first_line
    assert 'parts = Split(normalized, vbLf)' in first_line


def test_master_renderer_reports_precise_runtime_stage_on_failure():
    code = MASTER_RENDER_MODULE_CODE
    assert 'renderStage = "opening PATIENTS"' in code
    assert 'renderStage = "opening PATIENT_PLANNER"' in code
    assert 'renderStage = "opening MASTER_SCHEDULE"' in code
    assert 'renderStage = "locating patient rows"' in code
    assert 'renderStage = "locating MASTER row"' in code
    assert 'renderStage = "writing treatment cells"' in code
    assert '"MASTER refresh failed at " & renderStage & ": " & errorText' in code


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


def test_patient_program_runtime_gets_shared_validation_bridge():
    module = _fake_code_module(PATIENT_PROGRAM_MODULE_CODE.replace("\n", "\r\n"))

    _patch_runtime_validation(module)

    assert "Public Function ValidatePatientProgram" in module.text
    assert '"validate-vba "' in module.text


def test_patient_program_runtime_validation_patch_is_procedure_scoped():
    module = _fake_code_module(
        "Option Explicit\r\n"
        "\r\n"
        "Public Function CompletePatientProgram(ByVal patientId As String, ByVal requestArguments As String) As String\r\n"
        "    CompletePatientProgram = RunProgramBridge(\"different spacing is fine\")\r\n"
        "End Function\r\n"
        "\r\n"
        "Public Function SomeOtherFunction() As String\r\n"
        "    SomeOtherFunction = \"ok\"\r\n"
        "End Function\r\n"
    )

    _patch_runtime_validation(module)

    assert "Public Function ValidatePatientProgram" in module.text
    assert module.text.index("Public Function ValidatePatientProgram") < module.text.index(
        "Public Function SomeOtherFunction"
    )


def test_patient_program_runtime_validation_patch_is_idempotent():
    module = _fake_code_module(
        PATIENT_PROGRAM_MODULE_CODE.replace(
            "Public Function FindPatientDisplayName",
            "Public Function ValidatePatientProgram(ByVal patientId As String, ByVal requestArguments As String) As String\n"
            "    ValidatePatientProgram = \"already here\"\n"
            "End Function\n\n"
            "Public Function FindPatientDisplayName",
        ).replace("\n", "\r\n")
    )

    _patch_runtime_validation(module)

    assert module.add_count == 0


def test_patient_program_form_validates_conflicts_before_auto_and_before_writes():
    module = _fake_code_module(PATIENT_PROGRAM_FORM_CODE.replace("\n", "\r\n"))

    _patch_program_conflict_validation(module)

    code = module.text
    assert "Private Function CurrentConflictText() As String" in code
    assert "ValidatePatientProgram(mPatientId, requestArguments)" in code
    assert 'MsgBox conflictText & vbCrLf & vbCrLf &' in code
    assert '"Διόρθωσε τη σύγκρουση και ξαναπάτησε Αυτόματη συμπλήρωση."' in code
    assert "If Not ConfirmConflictOverride(conflictText) Then" in code
    assert 'lblStatus.Caption = "Η αποθήκευση ακυρώθηκε για να διορθωθεί η σύγκρουση."' in code
    assert code.index("conflictText = CurrentConflictText()", code.index("Private Sub SavePatientProgramAtomic")) < code.index(
        'Set ws = ThisWorkbook.Worksheets("PATIENT_PLANNER")',
        code.index("Private Sub SavePatientProgramAtomic"),
    )


def test_conflict_form_has_explicit_continue_and_return_choices():
    code = CONFLICT_FORM_CODE
    assert "Public ContinueSelected As Boolean" in code
    assert "Private Sub cmdContinue_Click()" in code
    assert "Private Sub cmdReturn_Click()" in code
    assert "ContinueSelected = True" in code
    assert "ContinueSelected = False" in code
    assert "Θέλετε να συνεχίσετε την αποθήκευση ή να επιστρέψετε για διόρθωση;" in code


def test_conflict_control_sizing_tolerates_vbe_width_height_rejection():
    class FakeControl:
        Caption = ""
        Left = 0
        Top = 0

        @property
        def Width(self):
            return 100

        @Width.setter
        def Width(self, value):
            raise RuntimeError("Property '<unknown>.Width' can not be set")

        @property
        def Height(self):
            return 20

        @Height.setter
        def Height(self, value):
            raise RuntimeError("Property '<unknown>.Height' can not be set")

    class FakeControls:
        def __init__(self):
            self.control = FakeControl()

        def __call__(self, name):
            return self.control

        def Add(self, prog_id, name, visible):
            return self.control

    class FakeDesigner:
        def __init__(self):
            self.Controls = FakeControls()

    control = _ensure_conflict_control(
        FakeDesigner(), "Forms.Label.1", "lblConflict", "", 18, 18, 570, 220
    )

    assert control.Caption == ""
    assert control.Left == 18
    assert control.Top == 18


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
