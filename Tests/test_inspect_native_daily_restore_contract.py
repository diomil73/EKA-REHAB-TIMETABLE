"""Synthetic-only safety regression for native Apply/Restore inspection.

Does not launch Excel or read patient data. Real Windows Excel acceptance is
independent and always requires a locally isolated demo copy.
"""
from types import SimpleNamespace

import pytest

from Python.tools import inspect_native_daily_restore_contract as tool


DAILY = """
Private Sub ApplyCharacterStrike()
End Sub
Private Function VerifyCharacterStrike() As Boolean
End Function
Sub Stage()
NormalizeMasterPatientStatusPresentation master
ApplyCharacterStrike master.Cells(row, statusCol), False
End Sub
"""
LEDGER = """
Sub Commit()
ApplyValidatedDailyInputPayload(patchPath)
wsDaily.Cells.Find(What:="*"
Dim changeId As String
End Sub
"""


class Module:
    def __init__(self, content):
        self.code = content
        self.CountOfLines = len(content.splitlines())

    def Lines(self, from_line, count):
        assert from_line == 1 and count == self.CountOfLines
        return self.code


class Project:
    def __init__(self, daily=DAILY, ledger=LEDGER):
        self.modules = {
            "modDailyInputAction": SimpleNamespace(CodeModule=Module(daily)),
            "modDailyChangeLedger": SimpleNamespace(CodeModule=Module(ledger)),
        }

    def VBComponents(self, name):
        return self.modules[name]


class Workbook:
    def __init__(self, *, readonly=True, daily=DAILY, ledger=LEDGER):
        self.ReadOnly = readonly
        self.VBProject = Project(daily=daily, ledger=ledger)
        self.close_calls = []

    def Close(self, **kwargs):
        self.close_calls.append(kwargs)


class App:
    def __init__(self, *, readonly=True, daily=DAILY, ledger=LEDGER):
        self.Workbooks = self
        self.book = Workbook(readonly=readonly, daily=daily, ledger=ledger)
        self.open_args = []
        self.closed = False
        self.AutomationSecurity = None

    def Open(self, path, **kwargs):
        assert self.AutomationSecurity == 3
        assert kwargs["ReadOnly"] is True
        assert kwargs["UpdateLinks"] == 0
        self.open_args.append(kwargs)
        return self.book

    def Quit(self):
        self.closed = True


def demo(tmp_path):
    file = tmp_path / "SYNTHETIC_DEMO.xlsm"
    file.write_bytes(b"fictional and harmless")
    return file


def test_native_contract_is_read_only_macro_disabled_and_source_unchanged(tmp_path):
    source = demo(tmp_path)
    original = source.read_bytes()
    app = App()
    checks = tool.inspect_workbook(source, excel_factory=lambda: app)
    assert len(checks) == 9
    assert all(checks.values())
    assert source.read_bytes() == original
    assert app.book.close_calls == [{"SaveChanges": False}]
    assert app.closed
    assert len(app.open_args) == 1
    assert app.AutomationSecurity == 3


def test_missing_vba_contract_returns_failure_facts_without_writing(tmp_path):
    source = demo(tmp_path)
    app = App(ledger=LEDGER.replace("ApplyValidatedDailyInputPayload(patchPath)", "REMOVED"))
    checks = tool.inspect_workbook(source, excel_factory=lambda: app)
    assert checks["ledger.canonical_apply"] is False
    assert not all(checks.values())
    assert app.book.close_calls == [{"SaveChanges": False}]
    assert source.exists()


def test_writable_open_refused_even_if_excel_ignores_requested_readonly(tmp_path):
    source = demo(tmp_path)
    app = App(readonly=False)
    with pytest.raises(tool.NativeContractInspectionError, match="read-only"):
        tool.inspect_workbook(source, excel_factory=lambda: app)
    assert app.closed
    assert app.book.close_calls == [{"SaveChanges": False}]


def test_com_exception_is_redacted_and_source_preserved(tmp_path):
    source = demo(tmp_path)
    secret = f"Patient fixture, private local path: {source}"
    app = App()
    app.book.VBProject.VBComponents = lambda name: (_ for _ in ()).throw(
        RuntimeError(secret)
    )
    with pytest.raises(tool.NativeContractInspectionError) as caught:
        tool.inspect_workbook(source, excel_factory=lambda: app)
    assert secret not in str(caught.value)
    assert str(source) not in str(caught.value)
    assert app.closed


def test_original_source_mutation_fails_and_does_not_get_approved(tmp_path):
    source = demo(tmp_path)
    app = App()
    def mutating_open(path, **kwargs):
        assert app.AutomationSecurity == 3
        source.write_bytes(b"CHANGED")
        return app.book
    app.Open = mutating_open
    with pytest.raises(tool.NativeContractInspectionError, match="Source changed"):
        tool.inspect_workbook(source, excel_factory=lambda: app)


def test_cli_redacts_log_content_workbook_path_and_private_exception(tmp_path, monkeypatch, capsys):
    source = demo(tmp_path)
    secret = "FICTIONAL PRIVATE PATIENT NAME"
    monkeypatch.setattr(tool, "_log_status", lambda name: "PRESENT")
    monkeypatch.setattr(tool, "inspect_workbook", lambda path: {
        "daily.character_apply": True,
        "ledger.canonical_apply": True,
    })
    assert tool.main(["--input", str(source)]) == 0
    displayed = capsys.readouterr().out
    assert "PRESENT" in displayed
    assert secret not in displayed
    assert str(source) not in displayed
    assert "patient" not in displayed.casefold()

    def failing_check(path):
        raise RuntimeError(secret + str(path))
    monkeypatch.setattr(tool, "inspect_workbook", failing_check)
    assert tool.main(["--input", str(source)]) == 2
    displayed = capsys.readouterr().out
    assert "SAFETY STOP" in displayed
    assert secret not in displayed
    assert str(source) not in displayed


def test_log_status_checks_only_existence_and_never_reads_file(tmp_path, monkeypatch):
    monkeypatch.setattr(tool.tempfile, "gettempdir", lambda: str(tmp_path))
    secret_log = tmp_path / "eka_daily_input_last_error.log"
    secret_log.write_text("SYNTHETIC PRIVATE NAME", encoding="utf-8")
    monkeypatch.setattr(type(secret_log), "read_text",
                        lambda *args, **kwargs: (_ for _ in ()).throw(
                            AssertionError("Log contents must never be read")
                        ))
    assert tool._log_status(secret_log.name) == "PRESENT"
    assert tool._log_status("no-file.log") == "MISSING"
