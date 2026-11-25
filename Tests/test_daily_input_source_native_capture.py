"""Synthetic tests for native literal source-sheet capture."""
from pathlib import Path

import pytest
from openpyxl.utils import get_column_letter

import rehab_excel.daily_input_source_native_capture as capture


REQUIRED = {
    "PATIENTS": {"A1": "PatientID", "A2": "001", "C2": "FICTIONAL"},
    "PATIENT_PLANNER": {"A1": "PatientID", "A2": 1},
    "SETTINGS": {"A1": "THERAPIST"},
    "DAILY_INPUT": {"A1": "ServiceDate"},
    "DAILY_CHANGE_LEDGER": {"A1": "LedgerID"},
}


class _Used:
    Row = 1
    Column = 1

    def __init__(self, rows, columns):
        self.Rows = type("Rows", (), {"Count": rows})()
        self.Columns = type("Columns", (), {"Count": columns})()


class _Cell:
    def __init__(self, value, formula=None):
        self.Value2 = value
        self.Formula = formula or ""
        self.HasFormula = bool(formula)


class _Sheet:
    def __init__(self, values, formulas=None):
        self.values = values
        self.formulas = formulas or {}
        self.UsedRange = _Used(
            max((int("".join(ch for ch in a if ch.isdigit())) for a in values), default=1),
            max((sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed("".join(ch for ch in a if ch.isalpha())))) for a in values), default=1),
        )

    def Cells(self, row, column):
        address = f"{get_column_letter(column)}{row}"
        return _Cell(self.values.get(address), self.formulas.get(address))


class _Book:
    ReadOnly = True

    def __init__(self, sheets):
        self.sheets = sheets
        self.closed = False

    def Worksheets(self, name):
        return self.sheets[name]

    def Close(self, SaveChanges=False):
        self.closed = True


class _App:
    Version = "16.0"
    AutomationSecurity = 0
    Visible = True
    DisplayAlerts = True
    EnableEvents = True
    AskToUpdateLinks = True

    def __init__(self, sheets, formulas=None):
        formulas = formulas or {}
        self.book = _Book({
            name: _Sheet(values, formulas.get(name, {}))
            for name, values in sheets.items()
        })
        self.Workbooks = type("Workbooks", (), {"Open": lambda _self, *a, **k: self.book})()
        self.quit = False

    def Quit(self):
        self.quit = True


class _Probe:
    def __init__(self, path, sheets, *, formulas=None, errors=None):
        self.sheet_names = tuple(sheets)
        self.sheets = sheets
        self.formulas = formulas or {}
        self.errors = errors or {}

    def formula_cells(self, name):
        return self.formulas.get(name, {})

    def error_cells(self, name):
        return self.errors.get(name, {})

    def typed_cell_values(self, name):
        return self.sheets[name]


def _source(tmp_path):
    path = tmp_path / "native-source.xlsm"
    path.write_bytes(b"synthetic package")
    return path


def _required_sheets():
    return {name: dict(values) for name, values in REQUIRED.items()}


def test_native_source_capture_compares_all_literal_cells_and_optional_sheets(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    sheets["PATIENT_EXITS"] = {"A1": "PatientID", "A2": "P1"}
    monkeypatch.setattr(capture, "WorkbookPackageProbe", lambda _path: _Probe(_path, sheets))
    app = _App(sheets)

    result = capture.verify_daily_input_source_native_capture_readonly(
        path, excel_factory=lambda: app,
    )

    assert result.workbook_sha256
    assert result.native_excel_version == "16.0"
    assert result.checked_sheet_count == 6
    assert result.checked_populated_cells == sum(len(x) for x in sheets.values())
    assert result.checked_formula_cells == 0
    assert app.book.closed is True
    assert app.quit is True
    assert app.AutomationSecurity == 3
    assert app.Visible is False
    assert app.EnableEvents is False


def test_native_source_capture_rejects_formula_before_excel(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    monkeypatch.setattr(
        capture, "WorkbookPackageProbe",
        lambda _path: _Probe(_path, sheets, formulas={"PATIENTS": {"A2": "NOW()"}}),
    )
    monkeypatch.setattr(capture, "sys", type("Sys", (), {"platform": "win32"})())
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: pytest.fail("Excel must not open"),
        )


def test_native_source_capture_rejects_stored_error_before_excel(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    monkeypatch.setattr(
        capture, "WorkbookPackageProbe",
        lambda _path: _Probe(_path, sheets, errors={"SETTINGS": {"B4": "#REF!"}}),
    )
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: pytest.fail("Excel must not open"),
        )


def test_native_source_capture_rejects_text_identifier_coercion(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    monkeypatch.setattr(capture, "WorkbookPackageProbe", lambda _path: _Probe(_path, sheets))
    native_sheets = _required_sheets()
    native_sheets["PATIENTS"]["A2"] = 1
    app = _App(native_sheets)
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: app,
        )
    assert app.book.closed is True
    assert app.quit is True


def test_native_source_capture_rejects_missing_required_sheet(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    del sheets["DAILY_CHANGE_LEDGER"]
    monkeypatch.setattr(capture, "WorkbookPackageProbe", lambda _path: _Probe(_path, sheets))
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: pytest.fail("Excel must not open"),
        )


def test_native_source_capture_rejects_source_mutation_during_excel_cleanup(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    monkeypatch.setattr(capture, "WorkbookPackageProbe", lambda _path: _Probe(_path, sheets))
    app = _App(sheets)
    original_close = app.book.Close

    def close_and_mutate(SaveChanges=False):
        original_close(SaveChanges=SaveChanges)
        path.write_bytes(path.read_bytes() + b"changed")

    app.book.Close = close_and_mutate
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: app,
        )


def test_native_source_capture_rejects_formula_reported_by_excel(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    monkeypatch.setattr(capture, "WorkbookPackageProbe", lambda _path: _Probe(_path, sheets))
    app = _App(sheets)
    original_cells = app.book.sheets["PATIENTS"].Cells

    def formula_cell(row, column):
        cell = original_cells(row, column)
        if (row, column) == (2, 1):
            cell.HasFormula = True
        return cell

    app.book.sheets["PATIENTS"].Cells = formula_cell
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: app,
        )
    assert app.book.closed is True
    assert app.quit is True


def test_native_source_capture_rejects_used_range_beyond_fixed_bound(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    monkeypatch.setattr(capture, "WorkbookPackageProbe", lambda _path: _Probe(_path, sheets))
    app = _App(sheets)
    app.book.sheets["PATIENTS"].UsedRange = _Used(10001, 3)
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: app,
        )
    assert app.book.closed is True
    assert app.quit is True


def test_native_source_capture_allows_only_direct_patient_planner_source_links(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    formulas = {"PATIENT_PLANNER": {"A2": "PATIENTS!A2"}}
    monkeypatch.setattr(
        capture, "WorkbookPackageProbe",
        lambda _path: _Probe(_path, sheets, formulas=formulas),
    )
    app = _App(sheets, formulas={"PATIENT_PLANNER": {"A2": "=PATIENTS!A2"}})

    result = capture.verify_daily_input_source_native_capture_readonly(
        path, excel_factory=lambda: app,
    )

    assert result.checked_formula_cells == 1
    assert result.checked_populated_cells == sum(len(x) for x in sheets.values())
    assert app.book.closed is True
    assert app.quit is True


def test_native_source_capture_rejects_unsupported_patient_planner_formula(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    formulas = {"PATIENT_PLANNER": {"A2": "OTHER_SHEET!A2"}}
    monkeypatch.setattr(
        capture, "WorkbookPackageProbe",
        lambda _path: _Probe(_path, sheets, formulas=formulas),
    )
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: pytest.fail("Excel must not open"),
        )


def test_native_source_capture_rejects_native_planner_formula_result_mismatch(tmp_path, monkeypatch):
    path = _source(tmp_path)
    sheets = _required_sheets()
    formulas = {"PATIENT_PLANNER": {"A2": "PATIENTS!A2"}}
    monkeypatch.setattr(
        capture, "WorkbookPackageProbe",
        lambda _path: _Probe(_path, sheets, formulas=formulas),
    )
    app = _App(sheets, formulas={"PATIENT_PLANNER": {"A2": "=PATIENTS!A2"}})
    app.book.sheets["PATIENT_PLANNER"].values["A2"] = "WRONG VALUE"
    with pytest.raises(capture.DailyInputSourceNativeCaptureError):
        capture.verify_daily_input_source_native_capture_readonly(
            path, excel_factory=lambda: app,
        )
