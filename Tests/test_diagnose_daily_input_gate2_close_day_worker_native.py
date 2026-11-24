"""Contract tests for the disposable native CloseDay guard probe."""
from pathlib import Path

from tools import diagnose_daily_input_gate2_close_day_worker_native as probe


def test_only_disposable_shim_invokes_private_close_day_writer():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert source.count("app.Run(") == 1
    assert "DIOMIL_TestInvokeCloseDayGuard" in source
    assert "PersistCloseDayPayload" in source
    worker_body = source.split("def run_locked_close_day_worker", 1)[1].split(
        "\ndef diagnose_native_close_day_guard", 1
    )[0]
    assert "CloseOperationalDay" not in worker_body
    assert "CloseOperationalDay UI wrapper was NOT invoked" in source


def test_close_day_worker_uses_missing_payload_and_hash_guards():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "MUST_NOT_EXIST.close-day" in source
    assert "file_sha256(copy) != before" in source
    assert "file_sha256(source) != source_before" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert "TemporaryDirectory(" in source
    assert "create_probe_copy(source, copy)" in source
    assert "component.CodeModule.InsertLines" in source


def test_close_day_module_readback_is_strict():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "_module_source_mismatch(CLOSE_DAY_MODULE_CODE, installed)" in source
    assert "Provisioned CloseDay module readback mismatch; " in source
    assert "canonicalize_vba_source(base)" in source


def test_nonwindows_stops_without_touching_source(tmp_path, monkeypatch, capsys):
    source = tmp_path / "fictional.xlsm"
    source.write_bytes(b"FICTIONAL")
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(source)]) == 2
    assert source.read_bytes() == b"FICTIONAL"
    assert "SAFETY STOP" in capsys.readouterr().out
