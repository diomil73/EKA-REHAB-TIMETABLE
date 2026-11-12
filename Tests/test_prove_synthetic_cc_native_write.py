"""Synthetic COM mock contract for real Excel save/reopen diagnostic probe."""
from pathlib import Path
from zipfile import ZipFile

import pytest

from tools import prove_synthetic_cc_native_write as proof


class Cell:
    def __init__(self, value=None):
        self.Value = value


class Sheet:
    def __init__(self, name, cell=None):
        self.Name = name
        self.cell = cell or Cell()

    def Range(self, address):
        assert address == "A1"
        return self.cell


class Sheets:
    def __init__(self, state):
        self.state = state

    def __iter__(self):
        return iter(Sheet(name) for name in self.state["sheets"])

    def Add(self):
        added = Sheet("_NEW")
        self.state["added"] = added
        return added

    def __call__(self, name):
        assert name in self.state["sheets"]
        return Sheet(name, Cell(self.state["marker"]))


class Book:
    def __init__(self, state, readonly):
        self.state = state
        self.ReadOnly = readonly
        self.Worksheets = Sheets(state)
        self.readonly = readonly

    def Save(self):
        assert not self.readonly
        added = self.state["added"]
        self.state["sheets"].add(added.Name)
        self.state["marker"] = added.Range("A1").Value
        self.state["saves"] += 1

    def Close(self, *, SaveChanges):
        assert SaveChanges is False
        self.state["closes"] += 1


class Workbooks:
    def __init__(self, state, app):
        self.state, self.app = state, app

    def Open(self, path, **kwargs):
        assert self.app.AutomationSecurity == 3
        assert kwargs["UpdateLinks"] == 0
        return Book(self.state, kwargs["ReadOnly"])


class App:
    def __init__(self, state):
        self.AutomationSecurity = None
        self.Workbooks = Workbooks(state, self)
        self.state = state

    def Quit(self):
        self.state["quits"] += 1


class Client:
    def __init__(self):
        self.state = {"sheets": {"MASTER_SCHEDULE"}, "marker": None,
                      "saves": 0, "closes": 0, "quits": 0}

    def DispatchEx(self, name):
        assert name == "Excel.Application"
        return App(self.state)


def test_synthetic_native_write_proof_only_modifies_new_copy(tmp_path, monkeypatch):
    monkeypatch.setattr(proof.sys, "platform", "win32")
    source = tmp_path / "UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm"
    with ZipFile(source, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"SYNTHETIC")
    before = source.read_bytes()
    target = tmp_path / "UNIFIED_H10_DEMO_THERAPIST_NATIVE_WRITE_A.xlsm"
    client = Client()
    result = proof.native_write_reopen_proof(
        source, target, excel_client=client
    )
    assert all(result.values())
    assert source.read_bytes() == before == target.read_bytes()
    assert client.state["marker"] == proof.PROBE_MARKER
    assert client.state["saves"] == 1
    assert client.state["closes"] == 2
    assert client.state["quits"] == 2


def test_synthetic_native_write_proof_refuses_existing_copy(tmp_path, monkeypatch):
    monkeypatch.setattr(proof.sys, "platform", "win32")
    source = tmp_path / "UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm"
    source.write_bytes(b"dummy")
    output = tmp_path / "UNIFIED_H10_DEMO_THERAPIST_NATIVE_WRITE_A.xlsm"
    output.write_bytes(b"DO NOT OVERWRITE")
    with pytest.raises(proof.NativeWriteProofError):
        proof.native_write_reopen_proof(source, output, excel_client=Client())
    assert output.read_bytes() == b"DO NOT OVERWRITE"


def test_synthetic_native_write_proof_refuses_nondemo_source(tmp_path, monkeypatch):
    monkeypatch.setattr(proof.sys, "platform", "win32")
    source = tmp_path / "MASTER.xlsm"
    source.write_bytes(b"PRODUCTION")
    output = tmp_path / "UNIFIED_H10_DEMO_THERAPIST_NATIVE_WRITE_A.xlsm"
    with pytest.raises(proof.NativeWriteProofError, match="Only isolated"):
        proof.native_write_reopen_proof(source, output, excel_client=Client())
    assert not output.exists()


def test_synthetic_native_write_proof_freezes_failed_new_copy(tmp_path, monkeypatch):
    monkeypatch.setattr(proof.sys, "platform", "win32")
    source = tmp_path / "UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm"
    with ZipFile(source, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"SYNTHETIC")
    before = source.read_bytes()
    target = tmp_path / "UNIFIED_H10_DEMO_THERAPIST_NATIVE_WRITE_A.xlsm"
    client = Client()

    class FailingApp:
        def DispatchEx(self, name):
            raise RuntimeError("synthetic startup failure")

    with pytest.raises(proof.NativeWriteProofError, match="native-save"):
        proof.native_write_reopen_proof(source, target, excel_client=FailingApp())
    assert target.read_bytes() == before == source.read_bytes()
