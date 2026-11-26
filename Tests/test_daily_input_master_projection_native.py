"""Synthetic fake-COM tests for native MASTER readback; no real Excel is run."""
from hashlib import sha256

import pytest
from openpyxl import Workbook
from openpyxl.utils.cell import column_index_from_string

from rehab_excel.daily_input_master_projection_native import (
    MasterProjectionNativeReadbackError,
    verify_master_projection_native_readonly,
)
from rehab_excel.master_projection import MasterProjectionRow
from rehab_excel.daily_input_master_projection_readback import SOURCE_COLUMNS


def _workbook(path, items, *, doctor=False, expanded=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "MASTER_SCHEDULE"
    ws["B3"], ws["C3"] = "Θάλαμος", "Ασθενής"
    status_col = 13 if expanded else 11
    if expanded:
        ws["K3"], ws["L3"], ws["M3"] = "ΨΥΧΟΛΟΓΟΙ", "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ", "Κατάσταση"
    else:
        ws["K3"] = "Κατάσταση"
    for item in items:
        for col, source in SOURCE_COLUMNS.items():
            ws.cell(item.target_row, col).value = f"=PATIENT_PLANNER!{source}{item.planner_row}"
        ws.cell(item.target_row, status_col).value = f"=PATIENT_PLANNER!E{item.planner_row}"
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
    def __init__(self, sheet, planner_sheet, *, readonly=True):
        self.ReadOnly = readonly
        self.sheet = sheet
        self.planner_sheet = planner_sheet
        self.closed = None

    def Worksheets(self, name):
        if name == "MASTER_SCHEDULE":
            return self.sheet
        if name == "PATIENT_PLANNER":
            return self.planner_sheet
        raise AssertionError(name)

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


def _expected(doctor=False, psychology_display=None):
    return (
        MasterProjectionRow(5, 2, "P-001", "A01", "FICTIONAL PATIENT 1",
                            False, "active", "DOCTOR X" if doctor else None,
                            psychology_display),
    )


def _fake_app(items, *, room="A01", identity=None, status="active",
              infectious=None, readonly=True, version="16.0",
              wrong_formula=False, formula_error=False,
              source_formula_column=None, source_value_mismatch=False,
              expanded=False, psychology_value=None, afternoon_value=None,
              psychology_formula=False):
    item = items[0]
    identity = identity if identity is not None else (
        f"{item.display_name}\n✚ {item.responsible_doctor}"
        if item.responsible_doctor else item.display_name
    )
    infectious_value = item.infectious if infectious is None else infectious
    source_values = {
        "D": infectious_value,
        "B": room,
        "C": item.display_name,
        "E": status,
        "F": "NATIVE F VALUE",
        "I": "NATIVE I VALUE",
        "N": "NATIVE N VALUE",
        "L": "NATIVE L VALUE",
        "P": "NATIVE P VALUE",
        "R": "NATIVE R VALUE",
        "T": "NATIVE T VALUE",
    }
    status_col = 13 if expanded else 11
    values = {
        (3, 2): Cell("Θάλαμος"),
        (3, 3): Cell("Ασθενής"),
    }
    if expanded:
        values.update({
            (3, 11): Cell("ΨΥΧΟΛΟΓΟΙ"),
            (3, 12): Cell("ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ"),
            (3, 13): Cell("Κατάσταση"),
        })
    else:
        values[(3, 11)] = Cell("Κατάσταση")
    for col, source in SOURCE_COLUMNS.items():
        value = identity if col == 3 else source_values[source]
        formula = None if col == 3 and item.responsible_doctor else (
            f"=PATIENT_PLANNER!{source}{item.planner_row}"
        )
        if wrong_formula and col == 2:
            formula = "=PATIENT_PLANNER!B99"
        if formula_error and col == 4:
            value = "#REF!"
        if source_value_mismatch and col == 4:
            value = "MISMATCHED FORMULA RESULT"
        values[(item.target_row, col)] = Cell(value, formula)
    values[(item.target_row, status_col)] = Cell(status, "=PATIENT_PLANNER!E2")
    if expanded:
        psychology = item.psychology_display if psychology_value is None else psychology_value
        values[(item.target_row, 11)] = Cell(
            psychology,
            "=PATIENT_PLANNER!K2" if psychology_formula else None,
        )
        values[(item.target_row, 12)] = Cell(afternoon_value)
    planner_values = {
        (2, column_index_from_string(source)): Cell(value)
        for source, value in source_values.items()
    }
    if source_formula_column:
        planner_values[(2, column_index_from_string(source_formula_column))] = Cell(
            source_values[source_formula_column],
            f"=OTHER_SHEET!A1",
        )
    book = Book(Sheet(values), Sheet(planner_values), readonly=readonly)
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


@pytest.mark.parametrize("failure", ["room", "identity", "status", "infectious", "formula", "formula_error", "source_formula", "source_value_mismatch", "readonly", "version"])
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
        "source_formula_column": "F" if failure == "source_formula" else None,
        "source_value_mismatch": failure == "source_value_mismatch",
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


def test_native_master_readback_supports_expanded_psychology_afternoon_schema(tmp_path):
    path = tmp_path / "synthetic-master-expanded.xlsm"
    expected = _expected(
        psychology_display="16:30\\nΔΕ, ΠΕ\\nΣΥΝΘΕΤΙΚΟΣ ΨΥΧΟΛΟΓΟΣ",
    )
    _workbook(path, expected, expanded=True)
    app, book = _fake_app(expected, expanded=True)

    result = verify_master_projection_native_readonly(
        path, expected_projection=expected, excel_factory=lambda: app,
    )

    assert result.status_column == 13
    assert result.projected_rows == 1
    assert result.formula_cells_checked == 11
    assert result.evaluated_cells_checked == 12
    assert app.quit is True
    assert book.closed is False

@pytest.mark.parametrize(
    "kwargs",
    [
        {"psychology_value": "WRONG PSYCHOLOGY"},
        {"afternoon_value": "UNAPPROVED PM DATA"},
        {"psychology_formula": True},
    ],
)
def test_native_master_readback_rejects_unverified_k_or_nonblank_l(tmp_path, kwargs):
    path = tmp_path / "synthetic-master-expanded-invalid.xlsm"
    expected = _expected(psychology_display="16:30\\nΔΕ\\nSYNTHETIC PROVIDER")
    _workbook(path, expected, expanded=True)
    app, book = _fake_app(expected, expanded=True, **kwargs)

    with pytest.raises(MasterProjectionNativeReadbackError):
        verify_master_projection_native_readonly(
            path, expected_projection=expected, excel_factory=lambda: app,
        )

    assert app.quit is True
    assert book.closed is False
