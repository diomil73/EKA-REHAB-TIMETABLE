"""Contract tests for the disposable-only native Restore guard probe."""
from pathlib import Path

from tools import diagnose_daily_input_gate2_restore_worker_native as probe


def test_only_disposable_restore_shim_invokes_private_worker():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert source.count("app.Run(") == 1
    assert "DIOMIL_TestInvokeRestoreGuard" in source
    assert "ApplyDailyChangeRestorePayload" in source
    assert "RestoreSelectedDailyChanges" not in source
    assert "RunDailyChangeRestore" not in source


def test_restore_worker_gate_is_probed_with_missing_payload_and_hashes():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "MUST_NOT_EXIST.restore" in source
    assert "file_sha256(copy) != before" in source
    assert "file_sha256(source) != source_before" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert "TemporaryDirectory(" in source
    assert "create_probe_copy(source, copy)" in source
    assert "component.CodeModule.InsertLines" in source


def test_nonwindows_stops_without_touching_source(tmp_path, monkeypatch, capsys):
    source = tmp_path / "fictional.xlsm"
    source.write_bytes(b"FICTIONAL")
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(source)]) == 2
    assert source.read_bytes() == b"FICTIONAL"
    assert "SAFETY STOP" in capsys.readouterr().out


def test_restore_module_mismatch_reports_bounded_source_diff():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "_module_source_mismatch(VBA_CODE, installed)" in source
    assert "Provisioned Restore module readback mismatch; " in source
    assert "_module_source_mismatch(VBA_CODE, base)" in source
