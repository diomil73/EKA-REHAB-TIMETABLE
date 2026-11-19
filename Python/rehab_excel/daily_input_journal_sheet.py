"""Native Excel worksheet contract for DAILY_INPUT_JOURNAL.

Only provisions or inspects the protected, veryHidden journal sheet. It never
clears an existing journal and does not append operational state. Appending
requires a later transaction-aware writer with retained rollback evidence.
"""
from __future__ import annotations

from .daily_input_journal import (
    JOURNAL_HEADERS,
    JOURNAL_SHEET,
    DailyInputJournalError,
    decode_operation_journal,
)


class DailyInputJournalSheetError(RuntimeError):
    pass


def _cell_value(ws, row: int, column: int):
    cell = ws.Cells(row, column)
    if getattr(cell, "HasFormula", False):
        raise DailyInputJournalSheetError("Formula in DAILY_INPUT_JOURNAL")
    value = cell.Value2
    if getattr(cell, "Text", "") == "#VALUE!":
        raise DailyInputJournalSheetError("Excel error in DAILY_INPUT_JOURNAL")
    return value


def read_native_journal_rows(ws) -> tuple[tuple, ...]:
    if str(ws.Name) != JOURNAL_SHEET:
        raise DailyInputJournalSheetError("Unexpected DAILY_INPUT journal worksheet")
    headers = tuple(_cell_value(ws, 1, col) for col in range(1, 6))
    if headers != JOURNAL_HEADERS:
        raise DailyInputJournalSheetError("Invalid DAILY_INPUT journal worksheet headers")
    last_row = max(int(ws.Cells(ws.Rows.Count, col).End(-4162).Row) for col in range(1, 6))  # xlUp
    rows = []
    for row in range(2, last_row + 1):
        values = tuple(_cell_value(ws, row, col) for col in range(1, 6))
        if not any(value is not None for value in values):
            raise DailyInputJournalSheetError("Blank row inside append-only journal")
        if any(value is None for value in values):
            raise DailyInputJournalSheetError("Partial row inside append-only journal")
        if type(values[1]) is not int:
            # Excel COM may expose integral numeric cells as float.
            if type(values[1]) is float and values[1].is_integer():
                values = (values[0], int(values[1]), *values[2:])
            else:
                raise DailyInputJournalSheetError("Invalid journal revision cell type")
        rows.append(values)
    try:
        decode_operation_journal(rows)
    except DailyInputJournalError as exc:
        raise DailyInputJournalSheetError("Native journal history failed integrity validation") from exc
    return tuple(rows)


def ensure_daily_input_journal_sheet(workbook):
    """Create an empty protected/veryHidden sheet, or validate without mutation."""
    # Enumerate explicitly instead of treating every COM lookup error as
    # "sheet missing". A disconnected/invalid Excel session must fail closed.
    ws = None
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        candidate = workbook.Worksheets(index)
        if str(candidate.Name) == JOURNAL_SHEET:
            ws = candidate
            break
    exists = ws is not None

    if not exists:
        ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        ws.Name = JOURNAL_SHEET
        for column, header in enumerate(JOURNAL_HEADERS, start=1):
            ws.Cells(1, column).Value2 = header
        ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
        ws.EnableSelection = -4142  # xlNoSelection
        ws.Visible = 2  # xlSheetVeryHidden
        workbook.Save()

    if str(ws.Name) != JOURNAL_SHEET:
        raise DailyInputJournalSheetError("Journal worksheet name mismatch")
    if int(ws.Visible) != 2:
        raise DailyInputJournalSheetError("DAILY_INPUT_JOURNAL must be veryHidden")
    if not bool(ws.ProtectContents):
        raise DailyInputJournalSheetError("DAILY_INPUT_JOURNAL must be protected")

    rows = read_native_journal_rows(ws)
    if not exists:
        if rows:
            raise DailyInputJournalSheetError("New DAILY_INPUT_JOURNAL must start empty")
        if tuple(_cell_value(ws, 1, col) for col in range(1, 6)) != JOURNAL_HEADERS:
            raise DailyInputJournalSheetError("New journal header readback mismatch")
    return ws
