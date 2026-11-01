from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.operator_visibility_patch_vba import (
    patch_daily_input_patient_status_visibility,
    patch_daily_input_rich_text_strike_contract,
    patch_daily_input_workload_visibility,
)


def _patched_daily_input() -> str:
    source = patch_daily_input_workload_visibility(DAILY_INPUT_MODULE_CODE)
    source = patch_daily_input_rich_text_strike_contract(source)
    return patch_daily_input_patient_status_visibility(source)


def test_master_non_active_overlay_accepts_only_null_strike_as_indeterminate():
    source = _patched_daily_input()

    assert 'If PayloadValue(parts(11)) <> "master_daily_non_active_overlay" Then Exit Function' in source
    assert 'strikeValue = cell.Font.Strikethrough' in source
    assert 'If Not IsNull(strikeValue) Then Exit Function' in source


def test_master_non_active_source_tag_still_requires_red_and_rejects_definite_false():
    source = _patched_daily_input()

    marker = '''        Case "master_daily_non_active_overlay"
            strikeValue = cell.Font.Strikethrough
            If IsNull(strikeValue) Then
                VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
            Else
                VerifySourceTagPresentation = _
                    (CBool(strikeValue) = True And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
            End If'''
    assert marker in source


def test_patient_status_overlay_remains_strict_no_strike():
    source = _patched_daily_input()

    assert 'VerifyCharacterStrike(cell, False)' in source
    assert 'NormalizeMasterPatientStatusPresentation master' in source
