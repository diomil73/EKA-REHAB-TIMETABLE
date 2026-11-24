"""Contract tests for the disposable native Replacement guard probe."""
from pathlib import Path

from tools import diagnose_daily_input_gate2_replacement_worker_native as probe


def test_only_disposable_replacement_shim_invokes_private_worker():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert source.count("app.Run(") == 1
    assert "DIOMIL_TestInvokeReplacementGuard" in source
    assert "PersistReplacementPayload" in source
    assert "ResolveDailyReplacementsInPlace" not in source
    assert "ShowDoubleBookingReview" not in source


def test_replacement_worker_uses_missing_payload_and_hash_guards():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "MUST_NOT_EXIST.replacement" in source
    assert "file_sha256(copy) != before" in source
    assert "file_sha256(source) != source_before" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert "TemporaryDirectory(" in source
    assert "create_probe_copy(source, copy)" in source
    assert "component.CodeModule.InsertLines" in source


def test_replacement_module_readback_mismatch_is_not_bypassed():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "_module_source_mismatch(GLOBAL_REPLACEMENT_MODULE_CODE, installed)" in source
    assert "Provisioned Replacement module readback mismatch; " in source
    assert "canonicalize_vba_source(base)" in source


def test_nonwindows_stops_without_touching_source(tmp_path, monkeypatch, capsys):
    source = tmp_path / "fictional.xlsm"
    source.write_bytes(b"FICTIONAL")
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(source)]) == 2
    assert source.read_bytes() == b"FICTIONAL"
    assert "SAFETY STOP" in capsys.readouterr().out
