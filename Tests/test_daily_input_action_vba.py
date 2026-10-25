import inspect

from rehab_excel.daily_input_action_vba import (
    DAILY_INPUT_MACRO_NAME,
    DAILY_INPUT_MODULE_CODE,
    DAILY_INPUT_MODULE_NAME,
    install_daily_input_action,
)


def _messages(source: str) -> str:
    return "\n".join(line for line in source.splitlines() if "MsgBox" in line).casefold()


def test_daily_input_action_component_names_are_stable():
    assert DAILY_INPUT_MODULE_NAME == "modDailyInputAction"
    assert DAILY_INPUT_MACRO_NAME == "ApplyDailyInputPreview"


def test_daily_input_action_saves_before_hidden_validation():
    save_pos = DAILY_INPUT_MODULE_CODE.index("ThisWorkbook.Save")
    run_pos = DAILY_INPUT_MODULE_CODE.index('CreateObject("WScript.Shell").Run(commandLine, 0, True)')
    assert save_pos < run_pos
    assert ".Exec(" not in DAILY_INPUT_MODULE_CODE


def test_daily_input_action_uses_in_place_write_plan():
    source = DAILY_INPUT_MODULE_CODE

    assert "prepare_daily_input_in_place.py" in source
    assert '" --input "' in source
    assert '" --payload "' in source
    assert '" --response "' in source
    assert "Public Function ApplyValidatedDailyInputPayload" in source
    assert "ApplyPatch" in source
    assert "VerifyPatch" in source
    assert "RestorePatchesFromRollback" in source
    assert "ThisWorkbook.SaveCopyAs rollbackPath" in source


def test_daily_input_action_has_no_close_reopen_commit_handoff():
    source = DAILY_INPUT_MODULE_CODE

    assert "ThisWorkbook.Close" not in source
    assert "Application.OnTime" not in source
    assert "authoritative_commit_worker_cli.py" not in source
    assert "StartAuthoritativeCommit" not in source
    assert "post_reopen_target" not in source
    assert "DAILY_INPUT_APPLIED_PREVIEW.xlsm" not in source


def test_daily_input_action_reports_operator_facing_results_only():
    source = DAILY_INPUT_MODULE_CODE
    messages = _messages(source)

    assert "Η ημερήσια κατάσταση δεν εφαρμόστηκε" in source
    assert "Η ημερήσια κατάσταση εφαρμόστηκε επιτυχώς." in source
    assert "Η ημερήσια κατάσταση και οι αναπληρώσεις εφαρμόστηκαν επιτυχώς." in source
    assert "backend" not in messages
    assert "worker" not in messages
    assert "preview" not in messages
    assert "hash" not in messages
    assert "κλείσει προσωρινά" not in messages
    assert "ανοίξει ξανά" not in messages


def test_daily_input_action_routes_to_global_replacement_workflow_only_when_needed():
    source = DAILY_INPUT_MODULE_CODE

    assert 'replacementsNeeded = JsonLongValue(responseText, "replacement_sessions_needed")' in source
    assert "If replacementsNeeded > 0 Then" in source
    assert "ResolveDailyReplacementsInPlace()" in source
    assert "ActivateReplacementScreen" not in source
    assert 'Worksheets("REPLACEMENTS").Activate' not in source
    assert "GoToMaster" in source


def test_daily_input_success_refreshes_operational_views_and_workload():
    source = DAILY_INPUT_MODULE_CODE

    assert "RefreshOperationalViews" in source
    assert "RenderTherapistWorkloadStrip" in source
    assert 'master.Shapes("ekaWorkload_Refresh")' in source
    assert "ApplyMasterView" in source
    assert 'Worksheets("THERAPIST_DAILY")' in source


def test_daily_input_installer_adds_guided_replacement_and_workload_modules():
    source = inspect.getsource(install_daily_input_action)

    assert "install_global_replacement_action(vbproject)" in source
    assert "install_double_booking_review(vbproject)" in source
    assert "WORKLOAD_MODULE_NAME" in source
    assert "WORKLOAD_MODULE_CODE" in source


def test_daily_input_patch_application_preserves_presentation_contract():
    source = DAILY_INPUT_MODULE_CODE

    assert "infectious_yellow" in source
    assert "robotic_pink" in source
    assert "outpatient_light_blue" in source
    assert "student_active_green" in source
    assert "robotic_orange" in source
    assert "ApplyTextRuns" in source
    assert "VerifyTextRuns" in source


def test_daily_input_mixed_font_cells_are_handled_without_cdbl_null_failure():
    source = DAILY_INPUT_MODULE_CODE

    assert "fontSizeValue = cell.Font.Size" in source
    assert "If IsNull(fontSizeValue) Then" in source
    assert "Len(textRuns) = 0" in source
    assert "If CDbl(cell.Font.Size) < CDbl(minFontToken)" not in source


def test_daily_input_native_failure_writes_hidden_temp_diagnostic_only():
    source = DAILY_INPUT_MODULE_CODE
    messages = _messages(source)

    assert "WriteDailyInputDiagnostic" in source
    assert "eka_daily_input_last_error.log" in source
    assert 'Print #handle, "phase=" & phase' in source
    assert 'Print #handle, "patch_index=" & CStr(patchIndex)' in source
    assert 'Print #handle, "sheet=" & sheetName' in source
    assert 'Print #handle, "cell=" & cellAddress' in source
    assert "eka_daily_input_last_error.log" not in messages
