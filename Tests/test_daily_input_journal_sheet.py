from types import SimpleNamespace

import pytest

from rehab_excel.daily_input_journal import JOURNAL_HEADERS, JOURNAL_SHEET, decode_operation_journal
from rehab_excel.daily_input_journal_sheet import (
    DailyInputJournalSheetError,
    ensure_daily_input_journal_sheet,
    read_native_journal_rows,
    append_native_operation_revision,
)
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase, advance_daily_input_operation,
)


class Cell:
    def __init__(self, sheet, row, col):
        self.sheet, self.row, self.col = sheet, row, col
        self.Value2 = None
        self.HasFormula = False
        self.Text = ""

    def ClearContents(self):
        self.Value2 = None

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

    def Unprotect(self):
        self.ProtectContents = False
        self.protected = False

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
        if isinstance(name, int):
            values = list(self.workbook.sheets.values())
            if name < 1 or name > len(values):
                raise KeyError(name)
            return values[name - 1]
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
        self.sheets = {"Sheet1": FakeSheet("Sheet1")}
        self.Worksheets = Sheets(self)
        self.saved = 0
        self.ReadOnly = False

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

def test_rejects_partial_trailing_row_even_when_operation_id_column_is_blank():
    ws = FakeSheet()
    for col, value in enumerate(JOURNAL_HEADERS, 1):
        ws.Write(1, col, value)
    ws.Write(2, 5, "orphan payload")
    with pytest.raises(DailyInputJournalSheetError, match="Partial row"):
        read_native_journal_rows(ws)


def _prepared_operation():
    return DailyInputOperation(
        operation_id="FICTIONAL-OPERATION",
        source_sha256="a" * 64,
        service_date="2026-11-20",
    )


def _workbook_with_empty_journal():
    wb = FakeWorkbook()
    ws = FakeSheet()
    for col, header in enumerate(JOURNAL_HEADERS, start=1):
        ws.Write(1, col, header)
    wb.sheets[JOURNAL_SHEET] = ws
    return wb, ws


def test_native_journal_stages_revisions_without_save_and_exact_retry_is_inert():
    wb, ws = _workbook_with_empty_journal()
    prepared = _prepared_operation()
    assert append_native_operation_revision(wb, prepared) is True
    assert ws.ProtectContents and ws.Visible == 2
    assert wb.saved == 0
    first = read_native_journal_rows(ws)
    assert len(first) == 1
    assert decode_operation_journal(first) == (prepared,)
    assert append_native_operation_revision(wb, prepared) is False
    assert read_native_journal_rows(ws) == first

    committed = advance_daily_input_operation(
        prepared, DailyInputOperationPhase.INPUT_AND_LEDGER_COMMITTED
    )
    failed = advance_daily_input_operation(
        committed, DailyInputOperationPhase.FAILED, failure_code="FICTIONAL_FAILURE"
    )
    rolled_back = advance_daily_input_operation(
        failed, DailyInputOperationPhase.ROLLED_BACK, rollback_verified=True
    )
    for revision in (committed, failed, rolled_back):
        assert append_native_operation_revision(wb, revision)
    assert decode_operation_journal(read_native_journal_rows(ws)) == (
        prepared, committed, failed, rolled_back,
    )
    assert ws.ProtectContents and wb.saved == 0


def test_native_journal_rejects_readonly_missing_sheet_and_illegal_transition():
    wb, ws = _workbook_with_empty_journal()
    prepared = _prepared_operation()
    wb.ReadOnly = True
    with pytest.raises(DailyInputJournalSheetError, match="writable"):
        append_native_operation_revision(wb, prepared)
    wb.ReadOnly = False
    committed = advance_daily_input_operation(
        prepared, DailyInputOperationPhase.INPUT_AND_LEDGER_COMMITTED
    )
    with pytest.raises(DailyInputJournalSheetError, match="transition rejected"):
        append_native_operation_revision(wb, committed)
    assert read_native_journal_rows(ws) == ()
    assert ws.ProtectContents and wb.saved == 0

    wb.sheets.pop(JOURNAL_SHEET)
    with pytest.raises(DailyInputJournalSheetError, match="unavailable"):
        append_native_operation_revision(wb, prepared)


def test_native_journal_partial_write_rollback_retains_previous_history():
    wb, ws = _workbook_with_empty_journal()
    prepared = _prepared_operation()
    assert append_native_operation_revision(wb, prepared)
    before = read_native_journal_rows(ws)
    committed = advance_daily_input_operation(
        prepared, DailyInputOperationPhase.INPUT_AND_LEDGER_COMMITTED
    )

    class FailOnceCell(Cell):
        def __init__(self, sheet, row, col):
            self.armed = False
            self._value = None
            super().__init__(sheet, row, col)
            self.armed = True

        @property
        def Value2(self):
            return self._value

        @Value2.setter
        def Value2(self, value):
            if self.armed and value is not None:
                self.armed = False
                raise RuntimeError("PRIVATE PATIENT VALUE")
            self._value = value

    ws.values[(3, 3)] = FailOnceCell(ws, 3, 3)
    with pytest.raises(DailyInputJournalSheetError, match="restored and verified") as caught:
        append_native_operation_revision(wb, committed)
    assert "PRIVATE" not in str(caught.value)
    assert read_native_journal_rows(ws) == before
    assert ws.ProtectContents and ws.Visible == 2 and wb.saved == 0


def test_native_journal_fails_closed_when_rollback_cannot_clear_partial_row():
    wb, ws = _workbook_with_empty_journal()
    prepared = _prepared_operation()

    class NeverClearCell(Cell):
        def ClearContents(self):
            raise RuntimeError("PRIVATE PATIENT CONTENT")

    ws.values[(2, 1)] = NeverClearCell(ws, 2, 1)

    class FailWriteCell(Cell):
        def __init__(self, sheet, row, col):
            self.armed = False
            self._value = None
            super().__init__(sheet, row, col)
            self.armed = True

        @property
        def Value2(self):
            return self._value

        @Value2.setter
        def Value2(self, value):
            if self.armed and value is not None:
                self.armed = False
                raise RuntimeError("PRIVATE WRITE FAILURE")
            self._value = value

    ws.values[(2, 2)] = FailWriteCell(ws, 2, 2)
    with pytest.raises(DailyInputJournalSheetError, match="rollback or protection unverified") as caught:
        append_native_operation_revision(wb, prepared)
    assert "PRIVATE" not in str(caught.value)
    assert wb.saved == 0


def test_native_journal_refuses_unprotected_or_visible_sheet():
    wb, ws = _workbook_with_empty_journal()
    for attr, value in (("ProtectContents", False), ("Visible", -1)):
        setattr(ws, attr, value)
        with pytest.raises(DailyInputJournalSheetError, match="protected and veryHidden"):
            append_native_operation_revision(wb, _prepared_operation())
        setattr(ws, attr, True if attr == "ProtectContents" else 2)
    assert wb.saved == 0
