from rehab_excel.patient_program_master_render_vba import (
    MASTER_RENDER_MODULE_CODE,
    MASTER_RENDER_MODULE_NAME,
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
