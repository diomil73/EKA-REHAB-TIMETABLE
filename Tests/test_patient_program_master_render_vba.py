from rehab_excel.patient_program_master_render_vba import (
    MASTER_RENDER_MODULE_CODE,
    MASTER_RENDER_MODULE_NAME,
    _patch_code_module,
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


def test_master_render_patcher_matches_vbe_crlf_code_with_lf_markers():
    class FakeCodeModule:
        def __init__(self, text):
            self.text = text
            self.added = None

        @property
        def CountOfLines(self):
            return len(self.text.splitlines())

        def Lines(self, start, count):
            return self.text

        def DeleteLines(self, start, count):
            self.text = ""

        def AddFromString(self, text):
            self.text = text
            self.added = text

    module = FakeCodeModule(
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
    class FakeCodeModule:
        def __init__(self, text):
            self.text = text
            self.add_count = 0

        @property
        def CountOfLines(self):
            return len(self.text.splitlines())

        def Lines(self, start, count):
            return self.text

        def DeleteLines(self, start, count):
            self.text = ""

        def AddFromString(self, text):
            self.text = text
            self.add_count += 1

    module = FakeCodeModule(
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
