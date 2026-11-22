from pathlib import Path
from zipfile import ZipFile

import pytest

from Python.tools import diagnose_vba_export_repeatability as probe
from rehab_excel.vba_semantic_audit import VbaComponentEvidence


class DummyWorkbook:
    ReadOnly = True

    def __init__(self):
        self.close_calls = []

    def Close(self, **kwargs):
        self.close_calls.append(kwargs)


class DummyWorkbooks:
    def __init__(self, book):
        self.book = book
        self.open_calls = []

    def Open(self, path, **kwargs):
        self.open_calls.append(kwargs)
        return self.book


class DummyExcel:
    def __init__(self):
        self.book = DummyWorkbook()
        self.Workbooks = DummyWorkbooks(self.book)
        self.quit_calls = 0

    def Quit(self):
        self.quit_calls += 1


def fake_source(path):
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"test-vba")


def evidence(resource):
    return (VbaComponentEvidence(
        name="internal",
        component_type=3,
        code_digest="source",
        export_digest="designer",
        companion_digest=resource,
    ),)


def test_read_only_repeatability_with_identical_exports(tmp_path, monkeypatch):
    source = tmp_path / "test.xlsm"
    fake_source(source)
    original = source.read_bytes()
    excel = DummyExcel()
    reads = iter([evidence("same"), evidence("same")])
    monkeypatch.setattr(probe, "audit_vba_components", lambda book: next(reads))

    result = probe.probe_vba_export_repeatability(source, excel_factory=lambda: excel)
    assert result == {"repeatable": True, "mismatch": "none", "components": 1}
    assert source.read_bytes() == original
    assert excel.Workbooks.open_calls == [{
        "UpdateLinks": 0, "ReadOnly": True,
        "AddToMru": False, "IgnoreReadOnlyRecommended": True,
    }]
    assert excel.book.close_calls == [{"SaveChanges": False}]
    assert excel.quit_calls == 1
    assert list(tmp_path.iterdir()) == [source]


def test_read_only_repeatability_reports_changed_form_resources(tmp_path, monkeypatch):
    source = tmp_path / "test.xlsm"
    fake_source(source)
    excel = DummyExcel()
    reads = iter([evidence("before"), evidence("after")])
    monkeypatch.setattr(probe, "audit_vba_components", lambda book: next(reads))

    result = probe.probe_vba_export_repeatability(source, excel_factory=lambda: excel)
    assert result["repeatable"] is False
    assert "resource=1" in result["mismatch"]
    assert "internal" not in result["mismatch"]
    assert excel.book.close_calls == [{"SaveChanges": False}]
    assert excel.quit_calls == 1


def test_read_only_repeatability_fail_closed_on_export_error(tmp_path, monkeypatch):
    source = tmp_path / "test.xlsm"
    fake_source(source)
    original = source.read_bytes()
    excel = DummyExcel()
    def fail(book):
        raise RuntimeError("synthetic readback failed")
    monkeypatch.setattr(probe, "audit_vba_components", fail)

    with pytest.raises(probe.RepeatabilityProbeError, match="first export"):
        probe.probe_vba_export_repeatability(source, excel_factory=lambda: excel)
    assert source.read_bytes() == original
    assert excel.book.close_calls == [{"SaveChanges": False}]
    assert excel.quit_calls == 1
