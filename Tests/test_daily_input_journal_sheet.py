from types import SimpleNamespace

import pytest

from rehab_excel.daily_input_journal import JOURNAL_HEADERS, JOURNAL_SHEET
from rehab_excel.daily_input_journal_sheet import (
    DailyInputJournalSheetError,
    ensure_daily_input_journal_sheet,
    read_native_journal_rows,
)


class Cell:
    def __init__(self, sheet, row, col):
        self.sheet, self.row, self.col = sheet, row, col
        self.Value2 = None
        self.HasFormula = False
        self.Text = ""

    def End(self, direction):
        populated = [r for (r, c), cell in self.sheet.values.items()
                     if c == self.col and cell.Value2 is not None]
        return EndCell(max(populated, default=1))


class EndCell:
    def __init__(self, row):
        self.Row = row


class Column:
    def __init__(self, sheet, col):
        self.sheet, self.col = sheet, col

    def End(self, direction):
        rows = [r for r in range(1, self.sheet.max_row + 1)
                if self.sheet.Cells(r, self.col).Value2 is not None]
        return EndCell(max(rows, default=1))


class Cells:
    def __init__(self, sheet):
        self.sheet = sheet
        self.Rows = SimpleNamespace(Count=1000)

    def __call__(self, row, col):
        return self.sheet.values.setdefault((row, col), Cell(self.sheet, row, col))

    def __getitem__(self, key):
        raise AssertionError("Unexpected indexed access")


class FakeSheet:
    def __init__(self, name=JOURNAL_SHEET):
        self.Name = name
        self.Visible = 2
        self.ProtectContents = True
        self.protected = True
        self.values = {}
        self.max_row = 1
        self.Cells = Cells(self)

    def __call__(self, row, col):
        return self.Cells(row, col)

    @property
    def Rows(self):
        return SimpleNamespace(Count=1000)

    def Protect(self, **kwargs):
        self.ProtectContents = True
        self.protected = True

    def Write(self, row, col, value):
        self.Cells(row, col).Value2 = value
        self.max_row = max(self.max_row, row)


class Sheets:
    def __init__(self, workbook):
        self.workbook = workbook

    def __call__(self, name):
        if name not in self.workbook.sheets:
            raise KeyError(name)
        return self.workbook.sheets[name]

    def Add(self, After=None):
        sheet = FakeSheet("SheetNew")
        self.workbook.sheets[sheet.Name] = sheet
        return sheet

    @property
    def Count(self):
        return len(self.workbook.sheets)


class FakeWorkbook:
    def __init__(self):
        self.sheets = {}
        self.Worksheets = Sheets(self)
        self.saved = 0

    def Save(self):
        self.saved += 1


def test_read_native_journal_requires_exact_schema_and_protection():
    ws = FakeSheet()
    for col, value in enumerate(JOURNAL_HEADERS, 1):
        ws.Write(1, col, value)
    assert read_native_journal_rows(ws) == ()
    ws.ProtectContents = False
    with pytest.raises(DailyInputJournalSheetError):
        ensure_daily_input_journal_sheet(FakeWorkbookWith(ws))


def FakeWorkbookWith(ws):
    wb = FakeWorkbook()
    wb.sheets[JOURNAL_SHEET] = ws
    return wb


def test_ensure_creates_protected_veryhidden_empty_sheet_and_saves():
    wb = FakeWorkbook()
    ws = ensure_daily_input_journal_sheet(wb)
    assert ws.Name == JOURNAL_SHEET
    assert ws.Visible == 2
    assert ws.ProtectContents
    assert wb.saved == 1
    assert tuple(ws.Cells(1, col).Value2 for col in range(1, 6)) == JOURNAL_HEADERS


def test_existing_journal_is_validated_without_clearing_or_saving():
    wb = FakeWorkbook()
    ws = FakeSheet()
    for col, value in enumerate(JOURNAL_HEADERS, 1):
        ws.Write(1, col, value)
    ws.Write(2, 1, "sentinel")
    wb.sheets[JOURNAL_SHEET] = ws
    with pytest.raises(DailyInputJournalSheetError):
        ensure_daily_input_journal_sheet(wb)
    assert ws.Cells(2, 1).Value2 == "sentinel"
    assert wb.saved == 0


def test_visible_or_unprotected_existing_journal_fails_closed():
    for attr, value in (("Visible", -1), ("ProtectContents", False)):
        wb = FakeWorkbook()
        ws = FakeSheet()
        for col, header in enumerate(JOURNAL_HEADERS, 1):
            ws.Write(1, col, header)
        setattr(ws, attr, value)
        wb.sheets[JOURNAL_SHEET] = ws
        with pytest.raises(DailyInputJournalSheetError):
            ensure_daily_input_journal_sheet(wb)
