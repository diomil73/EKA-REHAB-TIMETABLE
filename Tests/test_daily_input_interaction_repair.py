import inspect

from rehab_excel import daily_input_interaction_repair as repair
from tools import repair_daily_input_interaction as repair_tool


def test_repair_targets_only_known_daily_input_ranges():
    assert repair.THERAPIST_FIRST_ROW == 7
    assert repair.THERAPIST_LAST_ROW == 16
    assert repair.PATIENT_FIRST_ROW == 21
    assert repair.PATIENT_LAST_ROW == 40


def test_repair_restores_therapist_reason_dropdown_without_rebuilding_sheet():
    source = inspect.getsource(repair.repair_daily_input_interaction)
    assert "_assert_layout(ws)" in source
    assert "_ensure_reason_list(workbook)" in source
    assert "_set_reason_validation(ws)" in source
    assert "Cells.Clear" not in source
    assert "Cells.ClearContents" not in source


def test_repair_unlocks_only_input_ranges_for_edit_and_delete():
    source = inspect.getsource(repair.repair_daily_input_interaction)
    assert 'Range(f"A{THERAPIST_FIRST_ROW}:F{THERAPIST_LAST_ROW}").Locked = False' in source
    assert 'Range(f"A{PATIENT_FIRST_ROW}:E{PATIENT_LAST_ROW}").Locked = False' in source


def test_repair_rewires_only_exact_legacy_direct_global_macro():
    assert repair.LEGACY_DIRECT_REPLACEMENT_MACRO == "resolvedailyreplacementsinplace"
    assert repair.DAILY_INPUT_APPLY_MACRO == "ApplyDailyInputPreview"
    source = inspect.getsource(repair._rewire_legacy_buttons)
    assert "if current == LEGACY_DIRECT_REPLACEMENT_MACRO" in source
    assert "button.OnAction = DAILY_INPUT_APPLY_MACRO" in source
    assert "shape.OnAction = DAILY_INPUT_APPLY_MACRO" in source


def test_repair_is_fail_closed_on_unknown_layout_or_protection():
    layout_source = inspect.getsource(repair._assert_layout)
    repair_source = inspect.getsource(repair.repair_daily_input_interaction)
    assert "refusing to repair" in layout_source
    assert "remained protected" in repair_source


def test_cli_uses_rollback_and_verifies_native_readback():
    source = inspect.getsource(repair_tool.repair_workbook)
    assert "shutil.copy2(path, rollback)" in source
    assert "repair_daily_input_interaction(workbook)" in source
    assert "_verify(workbook)" in source
    assert "shutil.copy2(rollback, path)" in source
