"""Synthetic contract tests for native execution of the disabled gate predicate."""
from pathlib import Path

from tools import diagnose_daily_input_gate2_predicate_native as probe


def test_predicate_runner_has_one_allowlisted_vba_target():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert 'app.Run(f"\'{workbook.Name}\'!DailyInputJournalApplyEnabled")' in source
    assert "ApplyDailyInputPreview" not in source
    assert "ApplyValidatedDailyInputPayload" not in source
    assert source.count("app.Run(") == 1
    assert "RunDailyChangeRestore" not in source
    assert "type(result) is not bool" in source


def test_native_predicate_diagnostic_uses_disposable_copy_and_source_hash():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "TemporaryDirectory(" in source
    assert "create_probe_copy(source, copy)" in source
    assert "_provision_copy(copy" in source
    assert "file_sha256(source) != before" in source
    assert "workbook.Close(SaveChanges=False)" in source


def test_nonwindows_stops_before_excel(tmp_path, monkeypatch, capsys):
    filename = tmp_path / "fictional.xlsm"
    filename.write_bytes(b"FICTIONAL")
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(filename)]) == 2
    assert filename.read_bytes() == b"FICTIONAL"
    assert "SAFETY STOP" in capsys.readouterr().out
