from pathlib import Path
from zipfile import ZipFile

import pytest

from Python.tools import diagnose_excel_noedit_save_ole as control
from rehab_excel.vba_semantic_audit import VbaComponentEvidence


class FakeBook:
    ReadOnly = False
    def __init__(self):
        self.saves = 0
        self.closed = []

    def Save(self):
        self.saves += 1

    def Close(self, **kwargs):
        self.closed.append(kwargs)


class FakeExcel:
    def __init__(self):
        self.book = FakeBook()
        self.calls = []
        self.Workbooks = self
        self.quit = 0

    def Open(self, path, **kwargs):
        self.calls.append(kwargs)
        return self.book

    def Quit(self):
        self.quit += 1


def make_source(path):
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"synthetic-test-vba")


def test_noedit_save_only_touches_and_deletes_temporary_copy(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_source(source)
    initial = source.read_bytes()
    excel = FakeExcel()
    monkeypatch.setattr(control, "audit_vba_components", lambda book: (
        VbaComponentEvidence("testForm", 3, "a", "b", "c"),
    ))
    snapshot = (("testform", (("f", "a"), ("o", "b"))),)
    saved = (("testform", (("f", "changed"), ("o", "b"))),)
    reads = iter([snapshot, saved])
    monkeypatch.setattr(control, "form_storage_evidence", lambda path, forms: next(reads))
    result = control.probe_excel_save_without_edits(source, excel_factory=lambda: excel)
    assert result["forms"] == 1
    assert result["differences"]["streams_changed"] == 1
    assert result["categories"]["form_f_changed"] == 1
    assert excel.book.saves == 1
    assert excel.book.closed == [{"SaveChanges": False}]
    assert excel.calls[0]["ReadOnly"] is False
    assert excel.quit == 1
    assert source.read_bytes() == initial
    assert list(tmp_path.iterdir()) == [source]


def test_noedit_control_cleans_up_on_save_failure(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_source(source)
    initial = source.read_bytes()
    excel = FakeExcel()
    monkeypatch.setattr(control, "audit_vba_components", lambda book: (
        VbaComponentEvidence("testForm", 3, "a", "b", "c"),
    ))
    monkeypatch.setattr(control, "form_storage_evidence", lambda path, forms: (
        ("testform", (("f", "a"),)),))
    def fail():
        raise RuntimeError("synthetic failure")
    excel.book.Save = fail
    with pytest.raises(control.NativeNoEditSaveError, match="save disposable workbook"):
        control.probe_excel_save_without_edits(source, excel_factory=lambda: excel)
    assert excel.book.closed == [{"SaveChanges": False}]
    assert excel.quit == 1
    assert source.read_bytes() == initial
    assert list(tmp_path.iterdir()) == [source]


def test_noedit_control_rejects_missing_source(tmp_path):
    with pytest.raises(control.NativeNoEditSaveError, match="source required"):
        control.probe_excel_save_without_edits(tmp_path / "missing.xlsm", excel_factory=lambda: None)
