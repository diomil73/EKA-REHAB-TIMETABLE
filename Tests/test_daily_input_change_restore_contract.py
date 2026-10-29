from __future__ import annotations

import inspect

from rehab_excel.daily_input_change_ledger_vba import MODULE_CODE as LEDGER_CODE
from rehab_excel.daily_input_change_navigation_vba import MODULE_CODE as NAV_CODE
from rehab_excel.full_daily_projection import build_full_therapist_daily_plan
from tools.prepare_daily_input_restore_in_place import prepare_daily_input_restore_payload


def test_daily_input_change_ledger_uses_source_row_identity_and_in_place_restore():
    assert "DAILY_INPUT:" in LEDGER_CODE
    assert "sourceRow = CLng(sourceRowText)" in LEDGER_CODE
    assert 'Range("A" & CStr(sourceRow) & ":F" & CStr(sourceRow)).ClearContents' in LEDGER_CODE
    assert "ApplyValidatedDailyInputPayload(payloadPath)" in LEDGER_CODE
    assert "SaveCopyAs rollbackPath" in LEDGER_CODE
    assert "RestoreDailyInputTransaction" in LEDGER_CODE
    assert "ThisWorkbook.Close" not in LEDGER_CODE
    assert "authoritative_commit_worker" not in LEDGER_CODE.casefold()


def test_daily_input_restore_has_hidden_diagnostic_from_first_smoke():
    assert "eka_daily_input_restore_last_error.log" in LEDGER_CODE
    assert "WriteRestoreDiagnostic" in LEDGER_CODE


def test_daily_input_ledger_reuses_existing_daily_change_panel_safely():
    assert 'Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"' in LEDGER_CODE
    assert 'Application.Run "RefreshDailyChangeLedgerView"' in LEDGER_CODE
    assert 'shp.OnAction = "RestoreDailyInputChangeRow"' in LEDGER_CODE
    assert 'Shapes(SHAPE_PREFIX & "Select_" & CStr(row)).Delete' in LEDGER_CODE


def test_workload_navigation_syncs_daily_changes_before_operator_restore():
    assert "ShowTherapistWorkload" in NAV_CODE
    assert "SyncDailyInputChangesToLedger" in NAV_CODE
    assert "KeepApplicationShell" in NAV_CODE


def test_restore_preparation_validates_source_change_and_rebuilds_without_dependencies():
    source = inspect.getsource(prepare_daily_input_restore_payload)
    assert "selected_changes = tuple(" in source
    assert "if len(selected_changes) != 1:" in source
    assert 'selected_change.change_type == "therapist_absence"' in source
    assert "released_therapist_session_ids(" in source
    assert "dependent_active_replacement_ids(" in source
    assert "dependent_generated_suspension_ids(" in source
    assert "ignored_rows = tuple(sorted({source_row, *generated_suspension_rows}))" in source
    assert "replacement.replacement_id not in set(replacement_restore_ids)" in source
    assert "allow_empty=True" in source
    assert "build_full_therapist_daily_plan(" in source
    assert "build_full_master_daily_plan(" in source
    assert "file_sha256(source) != source_hash" in source


def test_restore_response_carries_only_verified_dependency_identities():
    source = inspect.getsource(prepare_daily_input_restore_payload)
    assert '"replacement_restore_ids": "|".join(replacement_restore_ids)' in source
    assert '"generated_suspension_rows": ",".join(' in source
    assert '"released_session_count": len(released_session_ids)' in source


def test_restore_vba_mutates_dependencies_before_projection_and_verifies_them():
    dependency_step = LEDGER_CODE.index("If Not ApplyRestoreDependencyMutations(responseText) Then")
    clear_source = LEDGER_CODE.index('Range("A" & CStr(sourceRow) & ":F" & CStr(sourceRow)).ClearContents')
    apply_projection = LEDGER_CODE.index("If Not ApplyValidatedDailyInputPayload(payloadPath) Then")
    assert dependency_step < clear_source < apply_projection
    assert 'UCase$(Trim$(CStr(wsFacts.Cells(found.Row, 10).Value))) <> STATUS_ACTIVE' in LEDGER_CODE
    assert 'wsFacts.Cells(found.Row, 10).Value = STATUS_RESTORED' in LEDGER_CODE
    assert 'Trim$(CStr(wsInput.Cells(rowNumber, 4).Value)) <> AUTO_SUSPENSION_STATUS' in LEDGER_CODE
    assert "AUTO_SUSPENSION_REASON" in LEDGER_CODE


def test_restore_transaction_rolls_back_daily_input_and_replacement_fact_mutations():
    assert 'backup.Worksheets(DAILY_INPUT_SHEET).Range("A1:F5000").Copy' in LEDGER_CODE
    assert 'Set backupFacts = backup.Worksheets(REPLACEMENT_FACT_SHEET)' in LEDGER_CODE
    assert 'backupFacts.Range("A1:J5000").Copy Destination:=currentFacts.Range("A1:J5000")' in LEDGER_CODE
    assert "currentFacts.Protect" in LEDGER_CODE


def test_full_daily_projection_clears_stale_grid_before_rebuilding():
    source = inspect.getsource(build_full_therapist_daily_plan)
    assert 'intent=WriteIntent.CLEAR' in source
    assert 'source_tag="full_daily_base_clear"' in source
    assert "build_patient_centric_preview_plan(" in source
