import inspect

from rehab_excel.close_day_action_vba import (
    CLOSE_DAY_MACRO_NAME,
    CLOSE_DAY_MODULE_CODE,
    CLOSE_DAY_MODULE_NAME,
)
from tools import build_registration_menu_preview


def _user_visible_message_source(source: str) -> str:
    return "\n".join(line for line in source.splitlines() if "MsgBox" in line)


def test_close_day_action_is_explicit_and_in_place():
    source = CLOSE_DAY_MODULE_CODE

    assert CLOSE_DAY_MODULE_NAME == "modCloseDayAction"
    assert CLOSE_DAY_MACRO_NAME == "CloseOperationalDay"
    assert "vbYesNo + vbDefaultButton2" in source
    assert "ThisWorkbook.Save" in source
    assert "prepare_close_day_in_place.py" in source
    assert "PersistCloseDayPayload" in source
    assert 'ThisWorkbook.Worksheets("MASTER_SCHEDULE").Activate' in source

    assert "ThisWorkbook.Close" not in source
    assert "Application.OnTime" not in source
    assert "authoritative_commit_worker_cli.py" not in source
    assert "post_reopen_target" not in source
    assert "StartCloseDayCommit" not in source
    assert "ScheduleDeferredCloseDayHandoff" not in source


def test_close_day_validation_runs_hidden_without_console_window():
    source = CLOSE_DAY_MODULE_CODE

    assert ".Exec(" not in source
    assert 'CreateObject("WScript.Shell").Run(commandLine, 0, True)' in source


def test_close_day_action_is_fail_closed_with_operator_facing_guidance():
    source = CLOSE_DAY_MODULE_CODE
    messages = _user_visible_message_source(source).casefold()

    assert "If exitCode <> 0" in source
    assert "Η ημέρα δεν έκλεισε." in source
    assert "Ελέγξτε ότι δεν υπάρχουν εκκρεμείς συνεδρίες" in source
    assert "επικοινωνήστε με τον διαχειριστή" in source
    assert "backend" not in messages
    assert "worker" not in messages
    assert "preview" not in messages
    assert "hash" not in messages
    assert "Η ημέρα έκλεισε επιτυχώς." in source
    assert "ανοίξει ξανά" not in messages


def test_close_day_second_click_reports_already_closed_as_normal_state():
    source = CLOSE_DAY_MODULE_CODE

    duplicate_check = 'InStr(1, responseText, "already finalized", vbTextCompare) > 0'
    assert duplicate_check in source
    assert "Η επιλεγμένη ημέρα είναι ήδη κλεισμένη." in source
    assert source.index(duplicate_check) < source.index("If exitCode <> 0")


def test_close_day_in_place_write_has_preflight_readback_and_rollback():
    source = CLOSE_DAY_MODULE_CODE

    assert "PreflightSnapshot" in source
    assert "PreflightTherapists" in source
    assert "PreflightEvent" in source
    assert "RowMatchesPayload" in source
    assert "RollbackHistoryWrite" in source
    assert "FINALIZED_DAY_SNAPSHOTS" in source
    assert "THERAPIST_DAY_HISTORY" in source
    assert "DAY_HISTORY_EVENTS" in source
    assert "xlSheetVeryHidden" in source
    assert "ThisWorkbook.Save" in source


def test_canonical_registration_build_installs_and_verifies_close_day_action():
    install_source = inspect.getsource(
        build_registration_menu_preview._install_close_day_action_in_workbook
    )
    main_source = inspect.getsource(build_registration_menu_preview.main)
    clean_source = inspect.getsource(build_registration_menu_preview._prepare_clean_vba_copy)

    assert "install_close_day_action(vbproject)" in install_source
    assert "CLOSE_DAY_MODULE_NAME" in install_source
    assert "Close-day action module was not confirmed" in install_source
    assert "_install_close_day_action_in_workbook(output)" in main_source
    assert "_remove_component_if_present(vbproject, CLOSE_DAY_MODULE_NAME)" in clean_source
