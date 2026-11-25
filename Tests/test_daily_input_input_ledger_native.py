"""Synthetic native COM facade tests; never open real Excel or a patient workbook."""
from datetime import date, time

import pytest
from openpyxl import Workbook, load_workbook

from rehab_core.models import Patient, Session
from rehab_excel.application_contract import DAILY_CHANGE_LEDGER_HEADERS
from rehab_excel.daily_input_reader import read_daily_input
from rehab_excel.daily_input_input_ledger_native import (
    InputLedgerNativeReadbackError, verify_input_ledger_native_readonly,
)

DAY = date(2026, 10, 26)


class Cell:
    def __init__(self, value, *, formula=False):
        self.Value2 = value
        self.HasFormula = formula


class Cells:
    def __init__(self, sheet):
        self.sheet = sheet

    def __call__(self, row, column):
        cell = self.sheet.cell(row, column)
        return Cell(cell.value, formula=cell.data_type == "f")


class Rows:
    def __init__(self, count):
        self.Count = count


class UsedRange:
    def __init__(self, sheet):
        self.Row = 1
        self.Rows = Rows(sheet.max_row)


class Sheet:
    def __init__(self, sheet):
        self.sheet = sheet
        self.Cells = Cells(sheet)
        self.UsedRange = UsedRange(sheet)
        self.Visible = 2 if sheet.sheet_state == "veryHidden" else -1
        self.ProtectContents = sheet.protection.sheet


class Sheets:
    def __init__(self, workbook):
        self.book = workbook

    def __call__(self, name):
        return Sheet(self.book[name])


class Book:
    def __init__(self, workbook, *, read_only=True, fail_close=False):
        self.ReadOnly = read_only
        self.Worksheets = Sheets(workbook)
        self.close_args = []
        self.fail_close = fail_close

    def Close(self, *, SaveChanges):
        self.close_args.append(SaveChanges)
        if self.fail_close:
            raise RuntimeError("synthetic close failure")


class Books:
    def __init__(self, book):
        self.book = book
        self.open_kwargs = None

    def Open(self, path, **kwargs):
        self.open_kwargs = kwargs
        return self.book


class Excel:
    Version = "SYNTHETIC-EXCEL-16"

    def __init__(self, book, *, fail_quit=False):
        self.Workbooks = Books(book)
        self.quit_count = 0
        self.fail_quit = fail_quit

    def Quit(self):
        self.quit_count += 1
        if self.fail_quit:
            raise RuntimeError("synthetic quit failure")


def fixture(tmp_path):
    source = tmp_path / "SYNTHETIC_INPUT_LEDGER.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "DAILY_INPUT"
    ws["A2"], ws["B2"] = "Ημερομηνία", "26/10/2026"
    for col, value in enumerate(("Θεραπευτής", "Όλη ημέρα", "Από", "Έως", "Αιτία", "Σχόλιο"), 1):
        ws.cell(6, col).value = value
    for col, value in enumerate(("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο"), 1):
        ws.cell(20, col).value = value
    for col, value in enumerate(
        ("ΔΟΚΙΜΑΣΤΙΚΟΣ ΑΣΘΕΝΗΣ", "ΟΧΙ", "08:30", "ΑΠΩΝ", "ΣΥΝΘΕΤΙΚΗ ΑΙΤΙΑ"), 1
    ):
        ws.cell(21, col).value = value
    ledger = wb.create_sheet("DAILY_CHANGE_LEDGER")
    ledger.append(DAILY_CHANGE_LEDGER_HEADERS)
    ledger.sheet_state = "veryHidden"
    ledger.protection.sheet = True
    wb.save(source)
    wb.close()
    patients = (Patient("FICTIONAL-P", "ΔΟΚΙΜΑΣΤΙΚΟΣ ΑΣΘΕΝΗΣ"),)
    sessions = (Session("FICTIONAL-S", "FICTIONAL-P", "FICTIONAL-T", DAY, time(8, 30)),)
    changes = read_daily_input(source, patients=patients, sessions=sessions).source_changes
    wb = load_workbook(source, keep_vba=False)
    for change in changes:
        wb["DAILY_CHANGE_LEDGER"].append((
            change.change_id, DAY.isoformat(), change.change_type,
            change.subject_id, change.display_name, "Before", "After",
            change.reason, change.source_row, "ACTIVE", None, None,
        ))
    wb.save(source)
    wb.close()
    return source, patients, sessions, changes


def native_factory(source, *, mutation=None, read_only=True, fail_close=False, fail_quit=False):
    wb = load_workbook(source, keep_vba=False)
    if mutation == "wrong-ledger":
        wb["DAILY_CHANGE_LEDGER"]["D2"] = "WRONG-IDENTITY"
    elif mutation == "wrong-input":
        wb["DAILY_INPUT"]["E21"] = "DIFFERENT REASON"
    elif mutation == "wrong-date":
        wb["DAILY_INPUT"]["B2"] = "27/10/2026"
    elif mutation == "unprotected-ledger":
        wb["DAILY_CHANGE_LEDGER"].protection.sheet = False
    elif mutation == "visible-ledger":
        wb["DAILY_CHANGE_LEDGER"].sheet_state = "visible"
    elif mutation == "formula-ledger":
        wb["DAILY_CHANGE_LEDGER"]["H2"] = "=1+1"
    elif mutation == "duplicate-ledger":
        ledger = wb["DAILY_CHANGE_LEDGER"]
        ledger.append(tuple(c.value for c in ledger[2]))
    book = Book(wb, read_only=read_only, fail_close=fail_close)
    excel = Excel(book, fail_quit=fail_quit)
    return excel, book, wb


def _verify(source, patients, sessions, changes, excel):
    return verify_input_ledger_native_readonly(
        source, service_date=DAY, validated_changes=changes,
        patients=patients, sessions=sessions, excel_factory=lambda: excel,
    )


def test_native_readonly_exact_match_preserves_bytes_and_closes_excel(tmp_path):
    source, patients, sessions, changes = fixture(tmp_path)
    original = source.read_bytes()
    excel, book, temp = native_factory(source)
    try:
        result = _verify(source, patients, sessions, changes, excel)
    finally:
        temp.close()
    assert result.native_excel_version == "SYNTHETIC-EXCEL-16"
    assert result.source_change_count == result.active_ledger_count == 1
    assert book.close_args == [False]
    assert excel.quit_count == 1
    assert excel.AutomationSecurity == 3
    assert excel.EnableEvents is False
    assert excel.DisplayAlerts is False
    assert excel.AskToUpdateLinks is False
    assert excel.Workbooks.open_kwargs["UpdateLinks"] == 0
    assert excel.Workbooks.open_kwargs["ReadOnly"] is True
    assert source.read_bytes() == original


@pytest.mark.parametrize("mutation", [
    "wrong-ledger", "wrong-input", "wrong-date", "unprotected-ledger",
    "visible-ledger", "formula-ledger", "duplicate-ledger",
])
def test_native_disagreement_fails_closed_without_workbook_writes(tmp_path, mutation):
    source, patients, sessions, changes = fixture(tmp_path)
    original = source.read_bytes()
    excel, book, temp = native_factory(source, mutation=mutation)
    try:
        with pytest.raises(InputLedgerNativeReadbackError):
            _verify(source, patients, sessions, changes, excel)
    finally:
        temp.close()
    assert book.close_args == [False]
    assert excel.quit_count == 1
    assert source.read_bytes() == original


@pytest.mark.parametrize("failure", ["not-readonly", "close", "quit"])
def test_native_lifecycle_failure_fails_closed(tmp_path, failure):
    source, patients, sessions, changes = fixture(tmp_path)
    excel, book, temp = native_factory(
        source, read_only=failure != "not-readonly",
        fail_close=failure == "close", fail_quit=failure == "quit",
    )
    try:
        with pytest.raises(InputLedgerNativeReadbackError):
            _verify(source, patients, sessions, changes, excel)
    finally:
        temp.close()
    assert book.close_args == [False]
    assert excel.quit_count == 1


def test_preflight_evidence_cannot_be_empty(tmp_path):
    source, patients, sessions, changes = fixture(tmp_path)
    excel, book, temp = native_factory(source)
    try:
        with pytest.raises(InputLedgerNativeReadbackError, match="preflight"):
            _verify(source, patients, sessions, (), excel)
    finally:
        temp.close()
    assert excel.quit_count == 0


def test_oo_xml_failure_prevents_native_excel_launch(tmp_path):
    source, patients, sessions, changes = fixture(tmp_path)
    wb = load_workbook(source, keep_vba=False)
    wb["DAILY_CHANGE_LEDGER"]["D2"] = "BAD"
    wb.save(source)
    wb.close()
    calls = []
    with pytest.raises(InputLedgerNativeReadbackError, match="OOXML"):
        verify_input_ledger_native_readonly(
            source, service_date=DAY, validated_changes=changes,
            patients=patients, sessions=sessions,
            excel_factory=lambda: calls.append("launched"),
        )
    assert not calls


def test_module_does_not_enable_apply():
    import inspect
    import rehab_excel.daily_input_input_ledger_native as module

    source = inspect.getsource(module)
    assert "DailyInputJournalApplyEnabled = True" not in source
    assert "ThisWorkbook.Save" not in source
