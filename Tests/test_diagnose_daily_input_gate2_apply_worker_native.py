"""Contract tests for the isolated native locked Apply worker probe."""
from pathlib import Path

from tools import diagnose_daily_input_gate2_apply_worker_native as probe


def test_only_direct_apply_worker_is_invoked():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert source.count("app.Run(") == 1
    assert "ApplyValidatedDailyInputPayload" in source
    assert "ApplyDailyInputPreview" not in source
    assert "RunDailyChangeRestore" not in source
    assert "CloseOperationalDay" not in source


def test_missing_payload_and_package_hash_guards_are_present():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "if missing_payload.exists()" in source
    assert "file_sha256(copy) != before" in source
    assert "file_sha256(source) != source_before" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert "TemporaryDirectory(" in source
    assert "create_probe_copy(source, copy)" in source


def test_nonwindows_stops_without_touching_source(tmp_path, monkeypatch, capsys):
    source = tmp_path / "fictional.xlsm"
    source.write_bytes(b"FICTIONAL")
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(source)]) == 2
    assert source.read_bytes() == b"FICTIONAL"
    assert "SAFETY STOP" in capsys.readouterr().out
