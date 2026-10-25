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


def test_restore_preparation_ignores_only_selected_source_row_and_allows_empty_day():
    source = inspect.getsource(prepare_daily_input_restore_payload)
    assert "ignore_rows=(source_row,)" in source
    assert "allow_empty=True" in source
    assert "build_full_therapist_daily_plan(" in source
    assert "build_full_master_daily_plan(" in source
    assert "file_sha256(source) != source_hash" in source


def test_full_daily_projection_clears_stale_grid_before_rebuilding():
    source = inspect.getsource(build_full_therapist_daily_plan)
    assert 'intent=WriteIntent.CLEAR' in source
    assert 'source_tag="full_daily_base_clear"' in source
    assert "build_patient_centric_preview_plan(" in source
