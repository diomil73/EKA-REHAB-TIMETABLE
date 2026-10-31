import inspect

from rehab_excel.daily_change_ledger_in_place_vba import (
    HEADERS,
    LEDGER_SHEET,
    MODULE_NAME,
    VBA_CODE,
    install_daily_change_ledger,
    patch_daily_input_for_ledger,
)
from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.operator_visibility_patch_vba import (
    patch_daily_input_patient_status_visibility,
    patch_daily_input_workload_visibility,
)
from tools import build_registration_menu_preview


def test_ledger_contract_has_stable_module_sheet_and_headers():
    assert MODULE_NAME == "modDailyChangeLedger"
    assert LEDGER_SHEET == "DAILY_CHANGE_LEDGER"
    assert HEADERS == (
        "ChangeID", "Date", "ChangeType", "SubjectID", "Subject", "FromText",
        "ToText", "Reason", "SourceRow", "Status", "RestoredAt", "DisplayRow",
    )


def test_ledger_ui_renders_change_history_and_all_restore_controls():
    for token in (
        '"Αλλαγές ημέρας"',
        '"Από"',
        '"Προς"',
        '"Αιτία"',
        '"Κατάσταση"',
        '"Επαναφορά"',
        '"Επαναφορά επιλεγμένων"',
        '"Επαναφορά όλων"',
        "Public Sub RestoreDailyChangeRow()",
        "Public Sub RestoreSelectedDailyChanges()",
        "Public Sub RestoreAllDailyChanges()",
    ):
        assert token in VBA_CODE


def test_restore_is_hidden_in_session_and_uses_atomic_multi_id_preparer():
    assert 'WScript.Shell").Run(commandLine, 0, True)' in VBA_CODE
    assert "prepare_daily_change_restore_in_place.py" in VBA_CODE
    assert 'args = args & " --change-id " & QuoteArg(changeId)' in VBA_CODE
    assert "ApplyDailyChangeRestorePayload" in VBA_CODE
    assert "ThisWorkbook.Close" not in VBA_CODE
    assert "Application.OnTime" not in VBA_CODE
    assert "authoritative_commit_worker_cli.py" not in VBA_CODE


def test_restore_payload_is_fail_closed_with_rollback_and_semantic_source_scope():
    assert 'PayloadValue(parts(1)) <> INPUT_SHEET' in VBA_CODE
    assert 'PayloadValue(parts(3)) <> "A"' in VBA_CODE
    assert 'PayloadValue(parts(4)) <> "E"' in VBA_CODE
    assert "ActiveLedgerChangeExists" in VBA_CODE
    assert "RestoreDailyChangeTransaction rollbackPath" in VBA_CODE
    assert 'RestoreOneSheetFromRollback rollbackPath, "THERAPIST_DAILY"' in VBA_CODE
    assert 'RestoreOneSheetFromRollback rollbackPath, "MASTER_SCHEDULE"' in VBA_CODE


def test_therapist_absence_restore_is_visible_but_not_falsely_enabled():
    assert 'changeType = "patient_absence" Or changeType = "session_cancellation"' in VBA_CODE
    assert '"Επαναφορά σύντομα"' in VBA_CODE


def test_daily_input_patch_accepts_change_rows_and_restores_ledger_on_failure():
    patched = patch_daily_input_for_ledger(DAILY_INPUT_MODULE_CODE)
    assert 'If CStr(parts(0)) = "CHANGE" Then' in patched
    assert "ValidateDailyChangePayload parts" in patched
    assert "PersistDailyChangePayload parts" in patched
    assert "RestoreDailyChangeLedgerFromRollback rollbackPath" in patched
    assert 'If CStr(parts(0)) <> "PATCH" Then GoTo RestoreNext' in patched
    assert "RefreshDailyChangeLedgerView" in patched


def test_ledger_patch_composes_after_operator_status_normalization():
    operator_patched = patch_daily_input_workload_visibility(DAILY_INPUT_MODULE_CODE)
    operator_patched = patch_daily_input_patient_status_visibility(operator_patched)
    patched = patch_daily_input_for_ledger(operator_patched)
    expected = (
        "    ApplyMasterView\n"
        "    NormalizeMasterPatientStatusPresentation master\n"
        "    RefreshDailyChangeLedgerView\n"
        "    RefreshOperationalViews = True"
    )
    assert expected in patched
    assert patched.count("RefreshDailyChangeLedgerView") == 1


def test_canonical_build_installs_ledger_after_operator_patch():
    source = inspect.getsource(build_registration_menu_preview)
    assert "install_daily_change_ledger" in source
    assert 'stage = "installing daily change ledger"' in source
    assert source.index("install_operator_visibility_patch(vbproject)") < source.index(
        "install_daily_change_ledger(vbproject, workbook)"
    )


def test_installer_ensures_sheet_and_patches_daily_input():
    source = inspect.getsource(install_daily_change_ledger)
    assert "ensure_daily_change_ledger_sheet(workbook)" in source
    assert 'vbproject.VBComponents("modDailyInputAction")' in source
    assert "patch_daily_input_for_ledger(current)" in source
