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
    assert '"Αλλαγές για την "' in VBA_CODE
    assert '"Επαναφορά"' in VBA_CODE
    assert '"Επαναφορά επιλεγμένων"' in VBA_CODE
    assert '"Επαναφορά όλων"' in VBA_CODE
    assert "Public Sub RestoreDailyChangeRow()" in VBA_CODE
    assert "Public Sub RestoreSelectedDailyChanges()" in VBA_CODE
    assert "Public Sub RestoreAllDailyChanges()" in VBA_CODE


def test_restore_actions_use_authoritative_backend_instead_of_cell_copy():
    assert 'ResolveAuthoritativeRestoreScript()' in VBA_CODE
    assert 'restore_daily_replacement_facts.py' in VBA_CODE
    assert ' --replacement-id ' in VBA_CODE
    assert ' --all-active' in VBA_CODE
    assert ' --response ' in VBA_CODE
    assert 'RunAuthoritativeRestore(' in VBA_CODE
    assert 'Cells(ledgerRow, 13).Copy Destination:=wsDaily.Range(originalCell)' not in VBA_CODE
    assert 'Cells(ledgerRow, 14).Copy Destination:=wsDaily.Range(effectiveCell)' not in VBA_CODE


def test_selected_restore_builds_one_atomic_backend_request():
    assert 'replacementArgs = replacementArgs & " --replacement-id " & QuoteArg(replacementId)' in VBA_CODE
    assert 'If selectedCount = 0 Then' in VBA_CODE
    assert 'Call RunAuthoritativeRestore(replacementArgs, False)' in VBA_CODE


def test_per_row_restore_passes_exact_immutable_replacement_id():
    assert 'replacementId = Trim$(CStr(wsLedger.Cells(rowIndex, 1).Value))' in VBA_CODE
    assert 'RunAuthoritativeRestore(" --replacement-id " & QuoteArg(replacementId), False)' in VBA_CODE


def test_restore_all_uses_backend_active_set_semantics():
    assert 'RunAuthoritativeRestore("", True)' in VBA_CODE
    assert 'commandLine = commandLine & " --all-active"' in VBA_CODE


def test_authoritative_restore_commit_closes_and_reopens_daily_sheet():
    assert 'ResolveAuthoritativeCommitWorkerScript()' in VBA_CODE
    assert 'authoritative_commit_worker_cli.py' in VBA_CODE
    assert '"post_reopen_target"' in VBA_CODE
    assert 'JsonEscape(DAILY_SHEET)' in VBA_CODE
    assert 'ThisWorkbook.Close SaveChanges:=True' in VBA_CODE


def test_ledger_refresh_keeps_restored_rows_visible_as_history():
    assert 'wsDaily.Cells(displayRow, 6).Value = "Επαναφέρθηκε"' in VBA_CODE
    assert 'wsDaily.Cells(displayRow, 6).Value = "Ενεργή αλλαγή"' in VBA_CODE


def test_ledger_refresh_preserves_readable_therapist_daily_time_axis():
    assert 'wsDaily.Columns("A").ColumnWidth = 8' in VBA_CODE
    assert 'wsDaily.Range("A2:A8,A12:A18").NumberFormat = "hh:mm"' in VBA_CODE
    assert 'wsDaily.Range("A2:A8,A12:A18").HorizontalAlignment = xlCenter' in VBA_CODE
    assert 'wsDaily.Columns("A").ColumnWidth = 4' not in VBA_CODE


def test_preview_removes_only_obsolete_refresh_form_button_surgically():
    source = inspect.getsource(_remove_legacy_refresh_button)
    assert "ws.Buttons()" in source
    assert 'name == "button 1"' in source
    assert 'caption == "ανανέωση".casefold()' in source
    assert 'macro == "refreshall"' in source
    assert "button.Delete()" in source
