from datetime import date, datetime, time

from rehab_excel.daily_change_ledger_vba import (
    DailyChangeLedgerRecord,
    LEDGER_SHEET,
    MODULE_NAME,
    VBA_CODE,
    _excel_date_value,
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


def test_restore_refuses_to_clobber_newer_changes_and_uses_snapshots():
    assert 'If currentOriginal <> expectedOriginal Then' in VBA_CODE
    assert 'If effectiveCell <> originalCell And currentEffective <> expectedEffective Then' in VBA_CODE
    assert 'Cells(ledgerRow, 13).Copy Destination:=wsDaily.Range(originalCell)' in VBA_CODE
    assert 'Cells(ledgerRow, 14).Copy Destination:=wsDaily.Range(effectiveCell)' in VBA_CODE
    assert 'Cells(ledgerRow, 12).Value = STATUS_RESTORED' in VBA_CODE


def test_restore_refreshes_master_patient_program_back_to_recurring_state():
    assert 'Call RefreshMasterPatientProgram(CStr(wsLedger.Cells(ledgerRow, 3).Value))' in VBA_CODE


def test_ledger_refresh_keeps_restored_rows_visible_as_history():
    assert 'wsDaily.Cells(displayRow, 6).Value = "Επαναφέρθηκε"' in VBA_CODE
    assert 'wsDaily.Cells(displayRow, 6).Value = "Ενεργή αλλαγή"' in VBA_CODE
