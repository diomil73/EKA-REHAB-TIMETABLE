from pathlib import Path
from zipfile import ZipFile

import pytest

from Python.tools import diagnose_excel_journal_vs_noedit_ole as probe
from rehab_excel.vba_semantic_audit import VbaComponentEvidence


class FakeBook:
    ReadOnly = False
    VBASigned = False

    def __init__(self):
        self.saves = 0
        self.closed = []
        self.Worksheets = FakeSheets()

    def Save(self):
        self.saves += 1

    def Close(self, **kwargs):
        self.closed.append(kwargs)


class FakeSheets:
    Count = 1
    def __call__(self, index):
        return type("Sheet", (), {"Name": "Sheet1"})()


class FakeApp:
    def __init__(self):
        self.book = FakeBook()
        self.Workbooks = self
        self.open_args = []
        self.quits = 0

    def Open(self, path, **kwargs):
        self.open_args.append(kwargs)
        return self.book

    def Quit(self):
        self.quits += 1


def make_xlsm(path):
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"fake-vba")


def evidence():
    return (VbaComponentEvidence("frmSynthetic", 3, "code", "export", "frx"),)


def test_two_separate_copies_produce_private_comparison_and_clean_up(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    original = source.read_bytes()
    apps = []
    def factory():
        app = FakeApp()
        apps.append(app)
        return app
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    calls = []
    def install(book):
        calls.append(book)
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", install)
    stream_sets = iter([
        (("frmsynthetic", (("f", "same"), ("o", "same"))),),
        (("frmsynthetic", (("f", "same"), ("o", "same"))),),
    ])
    monkeypatch.setattr(probe, "form_storage_evidence", lambda path, names: next(stream_sets))
    result = probe.compare_noedit_and_journal_saves(source, excel_factory=factory)
    assert len(apps) == 2
    assert apps[0].book.saves == 1
    assert apps[1].book.saves == 0
    assert len(calls) == 1
    assert all(a.book.closed == [{"SaveChanges": False}] for a in apps)
    assert all(a.quits == 1 for a in apps)
    assert all(a.open_args[0]["ReadOnly"] is False for a in apps)
    assert result["differences"]["streams_changed"] == 0
    assert source.read_bytes() == original
    assert list(tmp_path.iterdir()) == [source]


def test_comparison_detects_new_journal_stream_changes_without_approval(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", lambda wb: None)
    stream_sets = iter([
        (("frmsynthetic", (("f", "same"), ("o", "control"))),),
        (("frmsynthetic", (("f", "same"), ("o", "changed"))),),
    ])
    monkeypatch.setattr(probe, "form_storage_evidence", lambda path, names: next(stream_sets))
    result = probe.compare_noedit_and_journal_saves(source, excel_factory=FakeApp)
    assert result["differences"]["streams_changed"] == 1
    assert result["categories"]["form_o_changed"] == 1


def test_comparison_refuses_signed_project_and_cleans_up(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    apps = []
    def factory():
        app = FakeApp()
        app.book.VBASigned = True
        apps.append(app)
        return app
    with pytest.raises(probe.SaveComparisonError, match="signed VBA"):
        probe.compare_noedit_and_journal_saves(source, excel_factory=factory)
    assert len(apps) == 1
    assert apps[0].book.closed == [{"SaveChanges": False}]
    assert apps[0].quits == 1
    assert list(tmp_path.iterdir()) == [source]
