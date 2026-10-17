import inspect
from datetime import date, datetime, time

from rehab_excel.daily_change_ledger_vba import (
    DailyChangeLedgerRecord,
    LEDGER_SHEET,
    MODULE_NAME,
    VBA_CODE,
    _excel_date_value,
    _remove_legacy_refresh_button,
)
from rehab_excel.daily_restore_in_session_vba import build_in_session_ledger_vba


EFFECTIVE_VBA = build_in_session_ledger_vba(VBA_CODE)


def test_change_record_formats_human_readable_transition():
    record = DailyChangeLedgerRecord(
        change_id="chg-1",
        change_date=date(2026, 10, 17),
        patient_id="P1",
        patient_name="Ασθενής Χ",
        original_provider="Θεραπευτής Ψ",
        original_time=time(9, 0),
        effective_provider="Θεραπευτής Α",
        effective_time=time(10, 30),
        original_cell="C7",
        effective_cell="E10",
    )

    assert record.description == "Ασθενής Χ: Θεραπευτής Ψ 09:00 -> Θεραπευτής Α 10:30"


def test_ledger_date_is_converted_to_excel_compatible_datetime():
    assert _excel_date_value(date(2026, 10, 19)) == datetime(2026, 10, 19, 0, 0)


def test_ledger_uses_hidden_sheet_and_stable_module_name():
    assert LEDGER_SHEET == "DAILY_CHANGE_LEDGER"
    assert MODULE_NAME == "modDailyChangeLedger"


def test_vba_renders_daily_change_list_and_all_restore_controls():
    assert '"Αλλαγές για την "' in EFFECTIVE_VBA
    assert '"Επαναφορά"' in EFFECTIVE_VBA
    assert '"Επαναφορά επιλεγμένων"' in EFFECTIVE_VBA
    assert '"Επαναφορά όλων"' in EFFECTIVE_VBA
    assert "Public Sub RestoreDailyChangeRow()" in EFFECTIVE_VBA
    assert "Public Sub RestoreSelectedDailyChanges()" in EFFECTIVE_VBA
    assert "Public Sub RestoreAllDailyChanges()" in EFFECTIVE_VBA


def test_restore_actions_use_authoritative_backend_instead_of_cell_copy():
    assert 'ResolveAuthoritativeRestoreScript()' in EFFECTIVE_VBA
    assert 'restore_daily_replacement_facts.py' in EFFECTIVE_VBA
    assert ' --replacement-id ' in EFFECTIVE_VBA
    assert ' --all-active' in EFFECTIVE_VBA
    assert ' --response ' in EFFECTIVE_VBA
    assert 'RunAuthoritativeRestore(' in EFFECTIVE_VBA
    assert 'Cells(ledgerRow, 13).Copy Destination:=wsDaily.Range(originalCell)' not in EFFECTIVE_VBA
    assert 'Cells(ledgerRow, 14).Copy Destination:=wsDaily.Range(effectiveCell)' not in EFFECTIVE_VBA


def test_selected_restore_builds_one_atomic_backend_request():
    assert 'replacementArgs = replacementArgs & " --replacement-id " & QuoteArg(replacementId)' in EFFECTIVE_VBA
    assert 'If selectedCount = 0 Then' in EFFECTIVE_VBA
    assert 'Call RunAuthoritativeRestore(replacementArgs, False)' in EFFECTIVE_VBA


def test_per_row_restore_passes_exact_immutable_replacement_id():
    assert 'replacementId = Trim$(CStr(wsLedger.Cells(rowIndex, 1).Value))' in EFFECTIVE_VBA
    assert 'RunAuthoritativeRestore(" --replacement-id " & QuoteArg(replacementId), False)' in EFFECTIVE_VBA


def test_restore_all_uses_backend_active_set_semantics():
    assert 'RunAuthoritativeRestore("", True)' in EFFECTIVE_VBA
    assert 'commandLine = commandLine & " --all-active"' in EFFECTIVE_VBA


def test_authoritative_restore_runs_hidden_and_applies_inside_open_workbook():
    assert 'shell.Run(commandLine, 0, True)' in EFFECTIVE_VBA
    assert 'ApplyAuthoritativeRestorePreview previewPath' in EFFECTIVE_VBA
    assert 'ThisWorkbook.SaveCopyAs rollbackPath' in EFFECTIVE_VBA
    assert 'SyncRestoreSheetSet previewWorkbook, ThisWorkbook' in EFFECTIVE_VBA
    assert 'RefreshDailyChangeLedgerView' in EFFECTIVE_VBA
    assert 'χωρίς κλείσιμο του αρχείου' in EFFECTIVE_VBA


def test_primary_restore_path_forbids_terminal_and_close_reopen_worker():
    assert 'Set exec = shell.Exec(commandLine)' not in EFFECTIVE_VBA
    assert 'ThisWorkbook.Close SaveChanges:=True' not in EFFECTIVE_VBA
    assert 'StartAuthoritativeRestoreCommit' not in EFFECTIVE_VBA
    assert 'ResolveAuthoritativeCommitWorkerScript' not in EFFECTIVE_VBA
    assert 'authoritative_commit_worker_cli.py' not in EFFECTIVE_VBA


def test_in_session_apply_syncs_authoritative_operational_sheet_set():
    for sheet_name in (
        "DAILY_REPLACEMENT_FACTS",
        "DAILY_CHANGE_LEDGER",
        "THERAPIST_DAILY",
        "MASTER_SCHEDULE",
        "DAILY_STATISTICS",
    ):
        assert f'"{sheet_name}"' in EFFECTIVE_VBA


def test_ledger_refresh_keeps_restored_rows_visible_as_history():
    assert 'wsDaily.Cells(displayRow, 6).Value = "Επαναφέρθηκε"' in EFFECTIVE_VBA
    assert 'wsDaily.Cells(displayRow, 6).Value = "Ενεργή αλλαγή"' in EFFECTIVE_VBA


def test_ledger_refresh_preserves_readable_therapist_daily_time_axis():
    assert 'wsDaily.Columns("A").ColumnWidth = 8' in EFFECTIVE_VBA
    assert 'wsDaily.Range("A2:A8,A12:A18").NumberFormat = "hh:mm"' in EFFECTIVE_VBA
    assert 'wsDaily.Range("A2:A8,A12:A18").HorizontalAlignment = xlCenter' in EFFECTIVE_VBA
    assert 'wsDaily.Columns("A").ColumnWidth = 4' not in EFFECTIVE_VBA


def test_preview_removes_only_obsolete_refresh_form_button_surgically():
    source = inspect.getsource(_remove_legacy_refresh_button)
    assert "ws.Buttons()" in source
    assert 'name == "button 1"' in source
    assert 'caption == "ανανέωση".casefold()' in source
    assert 'macro == "refreshall"' in source
    assert "button.Delete()" in source
