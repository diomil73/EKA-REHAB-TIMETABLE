from rehab_excel.daily_input_action_vba import (
    DAILY_INPUT_MACRO_NAME,
    DAILY_INPUT_MODULE_CODE,
    DAILY_INPUT_MODULE_NAME,
)


def test_daily_input_action_component_names_are_stable():
    assert DAILY_INPUT_MODULE_NAME == "modDailyInputAction"
    assert DAILY_INPUT_MACRO_NAME == "ApplyDailyInputPreview"


def test_daily_input_action_saves_workbook_before_running_backend():
    save_pos = DAILY_INPUT_MODULE_CODE.index("ThisWorkbook.Save")
    run_pos = DAILY_INPUT_MODULE_CODE.index('Set exec = shell.Exec(commandLine)')
    assert save_pos < run_pos


def test_daily_input_action_calls_existing_apply_tool_with_safe_preview_output():
    assert "apply_daily_input.py" in DAILY_INPUT_MODULE_CODE
    assert '" --input "' in DAILY_INPUT_MODULE_CODE
    assert '" --output "' in DAILY_INPUT_MODULE_CODE
    assert '" --overwrite"' in DAILY_INPUT_MODULE_CODE
    assert "DAILY_INPUT_APPLIED_PREVIEW.xlsm" in DAILY_INPUT_MODULE_CODE


def test_daily_input_action_reports_backend_failure_and_success():
    assert "exec.ExitCode <> 0" in DAILY_INPUT_MODULE_CODE
    assert "Η εφαρμογή του DAILY_INPUT δεν ολοκληρώθηκε" in DAILY_INPUT_MODULE_CODE
    assert "Δημιουργήθηκε ασφαλές preview της ημερήσιας κατάστασης." in DAILY_INPUT_MODULE_CODE
