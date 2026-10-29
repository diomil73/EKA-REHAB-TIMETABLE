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


def test_repair_unlocks_input_ranges_then_relocks_authoritative_active_rows():
    source = inspect.getsource(repair._apply_input_editability)
    assert 'Range(f"A{THERAPIST_FIRST_ROW}:F{THERAPIST_LAST_ROW}").Locked = False' in source
    assert 'Range(f"A{PATIENT_FIRST_ROW}:E{PATIENT_LAST_ROW}").Locked = False' in source
    assert 'Range(f"A{source_row}:F{source_row}").Locked = True' in source
    assert 'Range(f"A{source_row}:E{source_row}").Locked = True' in source


def test_active_row_detection_uses_only_active_daily_input_ledger_metadata():
    source = inspect.getsource(repair.active_daily_input_rows)
    assert repair.LEDGER_SHEET == "DAILY_CHANGE_LEDGER"
    assert repair.LEDGER_STATUS_COLUMN == 12
    assert repair.LEDGER_META_COLUMN == 18
    assert repair.LEDGER_META_PREFIX == "DAILY_INPUT:"
    assert repair.LEDGER_STATUS_ACTIVE == "ACTIVE"
    assert "status != LEDGER_STATUS_ACTIVE" in source
    assert "_parse_daily_input_source_row" in source


def test_active_metadata_parser_rejects_rows_outside_known_input_ranges():
    assert repair._parse_daily_input_source_row("DAILY_INPUT:therapist_absence:7") == 7
    assert repair._parse_daily_input_source_row("DAILY_INPUT:patient_absence:21") == 21
    assert repair._parse_daily_input_source_row("DAILY_INPUT:therapist_absence:6") is None
    assert repair._parse_daily_input_source_row("DAILY_INPUT:patient_absence:41") is None
    assert repair._parse_daily_input_source_row("other:patient_absence:21") is None
    assert repair._parse_daily_input_source_row("DAILY_INPUT:patient_absence:not-a-row") is None


def test_repair_enables_sheet_protection_after_editability_is_applied():
    source = inspect.getsource(repair.repair_daily_input_interaction)
    assert "_apply_input_editability(workbook, ws)" in source
    assert "ws.Protect(UserInterfaceOnly=True)" in source
    assert source.index("_apply_input_editability(workbook, ws)") < source.index(
        "ws.Protect(UserInterfaceOnly=True)"
    )


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
    verify_source = inspect.getsource(repair_tool._verify)
    assert "shutil.copy2(path, rollback)" in source
    assert "repair_daily_input_interaction(workbook)" in source
    assert "_verify(workbook)" in source
    assert "shutil.copy2(rollback, path)" in source
    assert "active_daily_input_rows(workbook)" in verify_source
    assert "ACTIVE therapist input row remained editable" in verify_source
    assert "ACTIVE patient input row remained editable" in verify_source
    assert "Unapplied therapist input row remained locked" in verify_source
    assert "Unapplied patient input row remained locked" in verify_source
