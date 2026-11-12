"""Synthetic Windows COM proof: no macros, saves, clinical cell reads or writes."""
from pathlib import Path
from zipfile import ZipFile

import pytest

from tools.probe_native_excel_readonly import (
    NativeExcelProbeError,
    inspect_native_excel_readonly,
)


class FakeSheet:
    def __init__(self, name):
        self.Name = name


class FakeBook:
    def __init__(self, *, read_only=True, names=None):
        self.ReadOnly = read_only
        self.Worksheets = [
            FakeSheet(name) for name in (
                names if names is not None else
                ["MASTER_SCHEDULE", "DAILY_INPUT", "PATIENTS", "THERAPIST_DAILY"]
            )
        ]
        self.closes = []

    def Close(self, *, SaveChanges):
        self.closes.append(SaveChanges)


class FakeWorkbooks:
    def __init__(self, app, book):
        self.app = app
        self.book = book
        self.opens = []

    def Open(self, path, **kwargs):
        assert self.app.AutomationSecurity == 3
        assert self.app.EnableEvents is False
        assert self.app.DisplayAlerts is False
        self.opens.append(kwargs)
        return self.book


class FakeExcel:
    def __init__(self, book):
        self.AutomationSecurity = None
        self.EnableEvents = True
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AskToUpdateLinks = True
        self.Workbooks = FakeWorkbooks(self, book)
        self.quits = 0

    def Quit(self):
        self.quits += 1


def isolated_xlsm(tmp_path):
    path = tmp_path / "isolated.xlsm"
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"SYNTHETIC VBA BYTES")
    return path


def test_native_excel_probe_enforces_read_only_and_closes_without_save(tmp_path):
    path = isolated_xlsm(tmp_path)
    before = path.read_bytes()
    book = FakeBook()
    app = FakeExcel(book)
    result = inspect_native_excel_readonly(path, excel_factory=lambda: app)
    assert result["native_excel_readonly"] is True
    assert result["clinical_data_validated"] is False
    assert app.Workbooks.opens == [{
        "UpdateLinks": 0,
        "ReadOnly": True,
        "AddToMru": False,
        "IgnoreReadOnlyRecommended": True,
    }]
    assert book.closes == [False]
    assert app.quits == 1
    assert path.read_bytes() == before


def test_native_probe_rejects_writable_excel_session_without_saving(tmp_path):
    path = isolated_xlsm(tmp_path)
    book = FakeBook(read_only=False)
    app = FakeExcel(book)
    with pytest.raises(NativeExcelProbeError, match="read-only"):
        inspect_native_excel_readonly(path, excel_factory=lambda: app)
    assert book.closes == [False]
    assert app.quits == 1


def test_native_probe_rejects_missing_operational_sheet(tmp_path):
    path = isolated_xlsm(tmp_path)
    book = FakeBook(names=["PATIENTS", "MASTER_SCHEDULE"])
    app = FakeExcel(book)
    with pytest.raises(NativeExcelProbeError, match="operational sheets"):
        inspect_native_excel_readonly(path, excel_factory=lambda: app)
    assert book.closes == [False]


def test_native_probe_rejects_missing_vba_before_excel_opens(tmp_path):
    path = tmp_path / "isolated.xlsm"
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", b"SYNTHETIC")
    with pytest.raises(NativeExcelProbeError, match="VBA project"):
        inspect_native_excel_readonly(
            path, excel_factory=lambda: pytest.fail("Excel must not start")
        )


def test_native_probe_detects_external_change_during_read_only_open(tmp_path):
    path = isolated_xlsm(tmp_path)
    book = FakeBook()
    app = FakeExcel(book)
    base_open = app.Workbooks.Open

    def tamper(path_text, **kwargs):
        book = base_open(path_text, **kwargs)
        Path(path_text).write_bytes(b"SIMULATED EXTERNAL CHANGE")
        return book

    app.Workbooks.Open = tamper
    with pytest.raises(NativeExcelProbeError, match="Source changed"):
        inspect_native_excel_readonly(path, excel_factory=lambda: app)
    assert book.closes == [False]
    assert app.quits == 1


def test_native_probe_does_not_open_if_macro_security_cannot_be_set(tmp_path):
    path = isolated_xlsm(tmp_path)
    calls = []

    class UnsupportedSecurityApp:
        def __setattr__(self, name, value):
            if name == "AutomationSecurity":
                raise RuntimeError("COM security unavailable")
            super().__setattr__(name, value)

        def Quit(self):
            calls.append("quit")

        @property
        def Workbooks(self):
            pytest.fail("Workbook must never open with unknown macro security")

    with pytest.raises(RuntimeError, match="security unavailable"):
        inspect_native_excel_readonly(path, excel_factory=UnsupportedSecurityApp)
