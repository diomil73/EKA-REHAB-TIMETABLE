"""Synthetic fake-COM coverage for full native THERAPIST_DAILY grid readback."""
from datetime import time
from hashlib import sha256

import pytest
from openpyxl import Workbook

from rehab_excel.daily_input_therapist_projection_native import (
    TherapistDailyCellExpectation,
    TherapistDailyNativeReadbackError,
    verify_therapist_daily_native_readonly,
)
from rehab_excel.layout_contract import THERAPIST_DAILY_BLOCKS, THERAPIST_DAILY_TIMES, cell_address


def _expected():
    headers = tuple(f"THERAPIST {i}" for i in range(1, 10)) * 2
    cells = []
    for _header, first, last in THERAPIST_DAILY_BLOCKS:
        for row in range(first, last + 1):
            for col in range(2, 11):
                cells.append(TherapistDailyCellExpectation(cell_address(row, col), None))
    cells[0] = TherapistDailyCellExpectation("B2", "FICTIONAL PATIENT A")
    cells[-1] = TherapistDailyCellExpectation("J18", "FICTIONAL PATIENT B")
    return headers, tuple(cells)


class Cell:
    def __init__(self, value=None, *, formula=False):
        self.Value = value
        self.Value2 = value
        self.Formula = "=1+1" if formula else ""
        self.HasFormula = formula


class Sheet:
    def __init__(self, cells):
        self.cells = cells

    def Cells(self, row, col):
        return self.cells.get((row, col), Cell())


class Book:
    def __init__(self, sheet, *, readonly=True, close_error=False):
        self.ReadOnly = readonly
        self.sheet = sheet
        self.close_error = close_error
        self.closed = None

    def Worksheets(self, name):
        assert name == "THERAPIST_DAILY"
        return self.sheet

    def Close(self, SaveChanges=False):
        self.closed = SaveChanges
        if self.close_error:
            raise RuntimeError("injected close failure")


class Workbooks:
    def __init__(self, book):
        self.book = book
        self.open_args = None

    def Open(self, *args, **kwargs):
        self.open_args = (args, kwargs)
        return self.book


class App:
    def __init__(self, book, *, version="16.0", quit_error=False):
        self.AutomationSecurity = None
        self.Visible = None
        self.DisplayAlerts = None
        self.EnableEvents = None
        self.AskToUpdateLinks = None
        self.Version = version
        self.Workbooks = Workbooks(book)
        self.quit_error = quit_error
        self.quit = False

    def Quit(self):
        self.quit = True
        if self.quit_error:
            raise RuntimeError("injected quit failure")


def _fake_app(headers, expectations, *, wrong_cell=None, formula_cell=None,
              wrong_header=False, wrong_time=False, readonly=True,
              version="16.0", close_error=False, quit_error=False):
    values = {}
    header_index = 0
    for header_row, first, last in THERAPIST_DAILY_BLOCKS:
        for col in range(2, 11):
            value = headers[header_index]
            if wrong_header and header_index == 0:
                value = "WRONG PROVIDER"
            values[(header_row, col)] = Cell(value)
            header_index += 1
        for row, slot in zip(range(first, last + 1), THERAPIST_DAILY_TIMES, strict=True):
            numeric = (slot.hour * 60 + slot.minute) / 1440
            if wrong_time and row == first:
                numeric = (slot.hour * 60 + slot.minute + 1) / 1440
            values[(row, 1)] = Cell(numeric)

    for item in expectations:
        row = int("".join(ch for ch in item.cell if ch.isdigit()))
        col_letters = "".join(ch for ch in item.cell if ch.isalpha())
        col = 0
        for char in col_letters:
            col = col * 26 + ord(char) - ord("A") + 1
        value = "WRONG PATIENT" if item.cell == wrong_cell else item.value
        values[(row, col)] = Cell(value, formula=item.cell == formula_cell)

    book = Book(Sheet(values), readonly=readonly, close_error=close_error)
    app = App(book, version=version, quit_error=quit_error)
    return app, book


def _workbook(tmp_path):
    path = tmp_path / "synthetic-therapist.xlsm"
    wb = Workbook()
    wb.active.title = "THERAPIST_DAILY"
    wb.save(path)
    wb.close()
    return path


def test_native_therapist_full_grid_exact_match_and_readonly_cleanup(tmp_path):
    path = _workbook(tmp_path)
    headers, cells = _expected()
    original = path.read_bytes()
    app, book = _fake_app(headers, cells)

    result = verify_therapist_daily_native_readonly(
        path, expected_provider_headers=headers, expected_cells=cells,
        excel_factory=lambda: app,
    )

    assert result.checked_provider_headers == 18
    assert result.checked_grid_cells == 126
    assert result.occupied_grid_cells == 2
    assert result.native_excel_version == "16.0"
    assert result.workbook_sha256 == sha256(original).hexdigest()
    assert app.AutomationSecurity == 3
    assert app.Visible is False
    assert app.DisplayAlerts is False
    assert app.EnableEvents is False
    assert app.AskToUpdateLinks is False
    assert app.Workbooks.open_args[1]["ReadOnly"] is True
    assert book.closed is False
    assert app.quit is True
    assert path.read_bytes() == original


@pytest.mark.parametrize("failure", [
    "wrong-cell", "formula-cell", "wrong-header", "wrong-time",
    "readonly", "version", "close-cleanup", "quit-cleanup",
])
def test_native_therapist_full_grid_fails_closed(tmp_path, failure):
    path = _workbook(tmp_path)
    headers, cells = _expected()
    kwargs = {
        "wrong_cell": "B2" if failure == "wrong-cell" else None,
        "formula_cell": "B2" if failure == "formula-cell" else None,
        "wrong_header": failure == "wrong-header",
        "wrong_time": failure == "wrong-time",
        "readonly": failure != "readonly",
        "version": "" if failure == "version" else "16.0",
        "close_error": failure == "close-cleanup",
        "quit_error": failure == "quit-cleanup",
    }
    app, book = _fake_app(headers, cells, **kwargs)
    with pytest.raises(TherapistDailyNativeReadbackError):
        verify_therapist_daily_native_readonly(
            path, expected_provider_headers=headers, expected_cells=cells,
            excel_factory=lambda: app,
        )
    assert app.quit is True
    if failure != "version":
        assert book.closed is False


def test_native_therapist_requires_exact_full_grid_expectations(tmp_path):
    path = _workbook(tmp_path)
    headers, cells = _expected()
    app, _book = _fake_app(headers, cells)
    with pytest.raises(TherapistDailyNativeReadbackError):
        verify_therapist_daily_native_readonly(
            path, expected_provider_headers=headers, expected_cells=cells[:-1],
            excel_factory=lambda: app,
        )
    assert app.quit is False


def test_native_therapist_rejects_duplicate_provider_within_block(tmp_path):
    path = _workbook(tmp_path)
    headers, cells = _expected()
    bad = (headers[0], headers[0], *headers[2:])
    app, _book = _fake_app(headers, cells)
    with pytest.raises(TherapistDailyNativeReadbackError):
        verify_therapist_daily_native_readonly(
            path, expected_provider_headers=bad, expected_cells=cells,
            excel_factory=lambda: app,
        )
    assert app.quit is False


def test_native_therapist_accepts_blank_provider_slot_when_all_cells_blank(tmp_path):
    path = _workbook(tmp_path)
    headers, cells = _expected()
    headers = (*headers[:-1], "")
    cells = tuple(
        TherapistDailyCellExpectation(item.cell, None if item.cell.startswith("J") else item.value)
        for item in cells
    )
    app, book = _fake_app(headers, cells)
    result = verify_therapist_daily_native_readonly(
        path, expected_provider_headers=headers, expected_cells=cells,
        excel_factory=lambda: app,
    )
    assert result.checked_provider_headers == 18
    assert result.checked_grid_cells == 126
    assert result.occupied_grid_cells == 1
    assert book.closed is False


def test_native_therapist_rejects_occupancy_under_blank_provider_header(tmp_path):
    path = _workbook(tmp_path)
    headers, cells = _expected()
    headers = (*headers[:-1], "")
    app, _book = _fake_app(headers, cells)
    with pytest.raises(TherapistDailyNativeReadbackError):
        verify_therapist_daily_native_readonly(
            path, expected_provider_headers=headers, expected_cells=cells,
            excel_factory=lambda: app,
        )
