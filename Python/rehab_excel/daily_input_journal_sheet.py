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
    append_operation_journal,
)
from .daily_input_operation import DailyInputOperation


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


def append_native_operation_revision(workbook, operation: DailyInputOperation) -> bool:
    """Stage ONE validated revision in a pre-existing journal, without saving.

    Intended only inside a larger, rollback-aware Apply/Restore transaction.
    Returns True for a newly staged revision; False for an exact idempotent
    retry. Never creates a journal, opens a workbook, saves, runs VBA, or
    reports clinical success. The caller must independently verify durability
    across Save/Close/Reopen and restore operational state on any failure.
    """
    try:
        if bool(workbook.ReadOnly):
            raise DailyInputJournalSheetError("Journal append requires writable disposable workbook")
        ws = workbook.Worksheets(JOURNAL_SHEET)
    except DailyInputJournalSheetError:
        raise
    except Exception:
        raise DailyInputJournalSheetError("Existing journal sheet unavailable") from None

    if (str(ws.Name) != JOURNAL_SHEET or int(ws.Visible) != 2
            or not bool(ws.ProtectContents)):
        raise DailyInputJournalSheetError("Existing journal must be protected and veryHidden")

    before = read_native_journal_rows(ws)
    try:
        expected = append_operation_journal(before, operation)
    except DailyInputJournalError:
        raise DailyInputJournalSheetError("Journal operation transition rejected") from None
    if expected == before:
        return False  # No mutation or protection toggling on exact retry.

    row_index = len(before) + 2
    if row_index > int(ws.Rows.Count):
        raise DailyInputJournalSheetError("Journal worksheet row limit reached")
    if any(_cell_value(ws, row_index, col) is not None for col in range(1, 6)):
        raise DailyInputJournalSheetError("Journal append target is not empty")

    touched = False
    try:
        ws.Unprotect()
        if bool(ws.ProtectContents):
            raise DailyInputJournalSheetError("Journal could not be unprotected")
        for column, value in enumerate(expected[-1], start=1):
            # Mark touched before writing, since a setter may partially succeed.
            touched = True
            ws.Cells(row_index, column).Value2 = value
        ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
        if not bool(ws.ProtectContents):
            raise DailyInputJournalSheetError("Journal could not be reprotected")
        if read_native_journal_rows(ws) != expected:
            raise DailyInputJournalSheetError("Journal staged revision readback mismatch")
        return True
    except Exception:
        # A failed attempt must not leave a partial revision as usable history.
        # Only the new target row may be cleared; prior revisions are immutable.
        rollback_verified = False
        try:
            if touched:
                if bool(ws.ProtectContents):
                    ws.Unprotect()
                if bool(ws.ProtectContents):
                    raise DailyInputJournalSheetError("Journal rollback cannot unlock sheet")
                for column in range(1, 6):
                    ws.Cells(row_index, column).ClearContents()
            if not bool(ws.ProtectContents):
                ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
            rollback_verified = (
                bool(ws.ProtectContents)
                and int(ws.Visible) == 2
                and read_native_journal_rows(ws) == before
            )
        except Exception:
            rollback_verified = False
        if rollback_verified:
            raise DailyInputJournalSheetError(
                "Journal staging failed; previous history restored and verified"
            ) from None
        raise DailyInputJournalSheetError(
            "Journal staging failed; rollback or protection unverified"
        ) from None
