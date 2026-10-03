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


def test_daily_input_action_calls_existing_apply_tool_with_response_contract():
    assert "apply_daily_input.py" in DAILY_INPUT_MODULE_CODE
    assert '" --input "' in DAILY_INPUT_MODULE_CODE
    assert '" --output "' in DAILY_INPUT_MODULE_CODE
    assert '" --response "' in DAILY_INPUT_MODULE_CODE
    assert '" --overwrite"' in DAILY_INPUT_MODULE_CODE
    assert "DAILY_INPUT_APPLIED_PREVIEW.xlsm" in DAILY_INPUT_MODULE_CODE


def test_daily_input_action_reports_backend_failure_and_verified_save_message():
    assert "exec.ExitCode <> 0" in DAILY_INPUT_MODULE_CODE
    assert "Η εφαρμογή του DAILY_INPUT δεν ολοκληρώθηκε" in DAILY_INPUT_MODULE_CODE
    assert "Η ημερήσια κατάσταση επαληθεύτηκε και είναι έτοιμη για αποθήκευση." in DAILY_INPUT_MODULE_CODE


def test_daily_input_action_wires_authoritative_commit_worker():
    assert "StartAuthoritativeCommit" in DAILY_INPUT_MODULE_CODE
    assert "authoritative_commit_worker_cli.py" in DAILY_INPUT_MODULE_CODE
    assert 'JsonStringValue(responseText, "source_sha256_before")' in DAILY_INPUT_MODULE_CODE
    assert 'JsonStringValue(responseText, "output_path")' in DAILY_INPUT_MODULE_CODE
    assert 'q & "expected_source_sha256" & q' in DAILY_INPUT_MODULE_CODE
    assert 'q & "preview_path" & q' in DAILY_INPUT_MODULE_CODE
    assert 'q & "reopen" & q & ":true,"' in DAILY_INPUT_MODULE_CODE


def test_daily_input_action_starts_worker_async_then_closes_workbook():
    start_pos = DAILY_INPUT_MODULE_CODE.index("If Not StartAuthoritativeCommit(")
    close_pos = DAILY_INPUT_MODULE_CODE.index("ThisWorkbook.Close SaveChanges:=True")
    assert start_pos < close_pos
    assert 'CreateObject("WScript.Shell").Run commandLine, 0, False' in DAILY_INPUT_MODULE_CODE
