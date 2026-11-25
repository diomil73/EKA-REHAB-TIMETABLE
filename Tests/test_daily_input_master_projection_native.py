"""Synthetic fake-COM tests for native MASTER readback; no real Excel is run."""
from hashlib import sha256

import pytest
from openpyxl import Workbook

from rehab_excel.daily_input_master_projection_native import (
    MasterProjectionNativeReadbackError,
    verify_master_projection_native_readonly,
)
from rehab_excel.master_projection import MasterProjectionRow
from rehab_excel.daily_input_master_projection_readback import SOURCE_COLUMNS


def _workbook(path, items, *, doctor=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "MASTER_SCHEDULE"
    ws["B3"], ws["C3"], ws["K3"] = "Θάλαμος", "Ασθενής", "Κατάσταση"
    for item in items:
        for col, source in SOURCE_COLUMNS.items():
            ws.cell(item.target_row, col).value = f"=PATIENT_PLANNER!{source}{item.planner_row}"
        ws.cell(item.target_row, 11).value = f"=PATIENT_PLANNER!E{item.planner_row}"
        if item.responsible_doctor:
            ws.cell(item.target_row, 3).value = f"{item.display_name}\n✚ {item.responsible_doctor}"
    wb.save(path)
    wb.close()


class Cell:
    def __init__(self, value=None, formula=None):
        self.Value = value
        self.Value2 = value
        self.Formula = formula or (value if isinstance(value, str) and value.startswith("=") else "")
        self.HasFormula = bool(self.Formula)


class Sheet:
    def __init__(self, values):
        self.values = values

    def Cells(self, row, col):
        return self.values.get((row, col), Cell())


class Book:
    def __init__(self, sheet, *, readonly=True):
        self.ReadOnly = readonly
        self.sheet = sheet
        self.closed = None

    def Worksheets(self, name):
        assert name == "MASTER_SCHEDULE"
        return self.sheet

    def Close(self, SaveChanges=False):
        self.closed = SaveChanges


class Workbooks:
    def __init__(self, book):
        self.book = book
        self.open_args = None

    def Open(self, *args, **kwargs):
        self.open_args = (args, kwargs)
        return self.book


class App:
    def __init__(self, book, *, version="16.0"):
        self.AutomationSecurity = None
        self.Visible = None
        self.DisplayAlerts = None
        self.EnableEvents = None
        self.AskToUpdateLinks = None
        self.Version = version
        self.Workbooks = Workbooks(book)
        self.quit = False

    def Quit(self):
        self.quit = True


def _expected(doctor=False):
    return (
        MasterProjectionRow(5, 2, "P-001", "A01", "FICTIONAL PATIENT 1",
                            False, "active", "DOCTOR X" if doctor else None),
    )


def _fake_app(items, *, room="A01", identity=None, status="active",
              infectious=None, readonly=True, version="16.0",
              wrong_formula=False, formula_error=False):
    item = items[0]
    identity = identity if identity is not None else (
        f"{item.display_name}\n✚ {item.responsible_doctor}"
        if item.responsible_doctor else item.display_name
    )
    values = {
        (3, 2): Cell("Θάλαμος"),
        (3, 3): Cell("Ασθενής"),
        (3, 11): Cell("Κατάσταση"),
    }
    for col, source in SOURCE_COLUMNS.items():
        value = {
            1: item.infectious if infectious is None else infectious,
            2: room,
            3: identity,
        }.get(col, "NATIVE FORMULA VALUE")
        formula = None if col == 3 and item.responsible_doctor else (
            f"=PATIENT_PLANNER!{source}{item.planner_row}"
        )
        if wrong_formula and col == 2:
            formula = "=PATIENT_PLANNER!B99"
        if formula_error and col == 4:
            value = "#REF!"
        values[(item.target_row, col)] = Cell(value, formula)
    values[(item.target_row, 11)] = Cell(status, "=PATIENT_PLANNER!E2")
    book = Book(Sheet(values), readonly=readonly)
    app = App(book, version=version)
    return app, book

def test_native_master_readback_checks_values_and_closes_without_saving(tmp_path):
    path = tmp_path / "synthetic-master.xlsm"
    expected = _expected()
    _workbook(path, expected)
    original = path.read_bytes()
    app, book = _fake_app(expected)

    result = verify_master_projection_native_readonly(
        path, expected_projection=expected, excel_factory=lambda: app,
    )

    assert result.projected_rows == 1
    assert result.header_row == 3
    assert result.status_column == 11
    assert result.native_excel_version == "16.0"
    assert result.formula_cells_checked == 11
    assert result.evaluated_cells_checked == 11
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


def test_native_master_readback_checks_literal_doctor_identity(tmp_path):
    path = tmp_path / "synthetic-master-doctor.xlsm"
    expected = _expected(doctor=True)
    _workbook(path, expected)
    app, _book = _fake_app(expected)
    result = verify_master_projection_native_readonly(
        path, expected_projection=expected, excel_factory=lambda: app,
    )
    assert result.projected_rows == 1
    assert result.formula_cells_checked == 10
    assert result.evaluated_cells_checked == 11
    assert app.quit is True


@pytest.mark.parametrize("failure", ["room", "identity", "status", "infectious", "formula", "formula_error", "readonly", "version"])
def test_native_master_readback_fails_closed_and_cleans_up(tmp_path, failure):
    path = tmp_path / "synthetic-master.xlsm"
    expected = _expected()
    _workbook(path, expected)
    kwargs = {
        "room": "WRONG ROOM" if failure == "room" else "A01",
        "identity": "WRONG PATIENT" if failure == "identity" else None,
        "status": "inactive" if failure == "status" else "active",
        "infectious": "YES" if failure == "infectious" else None,
        "wrong_formula": failure == "formula",
        "formula_error": failure == "formula_error",
        "readonly": failure != "readonly",
        "version": "" if failure == "version" else "16.0",
    }
    app, book = _fake_app(expected, **kwargs)
    with pytest.raises(MasterProjectionNativeReadbackError):
        verify_master_projection_native_readonly(
            path, expected_projection=expected, excel_factory=lambda: app,
        )
    assert app.quit is True
    if failure != "version":
        assert book.closed is False


def test_native_master_readback_rejects_mutable_or_wrong_expectations(tmp_path):
    path = tmp_path / "synthetic-master.xlsm"
    expected = _expected()
    _workbook(path, expected)
    app, _book = _fake_app(expected)
    with pytest.raises(MasterProjectionNativeReadbackError):
        verify_master_projection_native_readonly(
            path, expected_projection=list(expected), excel_factory=lambda: app,
        )
    assert app.quit is False
