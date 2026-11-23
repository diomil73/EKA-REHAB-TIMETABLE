"""Safety tests for the new native-only synthetic rollback diagnostic."""
from pathlib import Path

from tools import diagnose_daily_input_full_rollback_native as probe


def test_nonwindows_probe_refuses_before_reading_or_touching_source(tmp_path, monkeypatch, capsys):
    source = tmp_path / "FICTIONAL-DEMO.xlsm"
    source.write_bytes(b"FICTIONAL_SOURCE_UNTOUCHED")
    before = source.read_bytes()
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(source)]) == 2
    assert "requires Windows + Excel" in capsys.readouterr().out
    assert source.read_bytes() == before


def test_native_diagnostic_touches_only_disposable_copies_and_never_invokes_apply():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "TemporaryDirectory(" in source
    assert "create_probe_copy(source, working)" in source
    assert "create_probe_copy(working, snapshot)" in source
    assert "create_probe_copy(working, failed)" in source
    assert "file_sha256(source) != source_sha" in source
    assert "_run_vba_hook(" in source
    assert "append_native_operation_revision(workbook, next_state)" in source
    assert "build_rollback_candidate(" in source
    assert "audit_restored_business_state_before_rollback(" in source
    assert "audit_full_workbook_rollback(" in source
    assert "verify_native_excel_readonly(candidate, failed_rows)" in source
    assert "verify_native_excel_readonly(terminal, terminal_rows)" in source
    assert "ApplyDailyInputPreview" not in source
    assert "ApplyValidatedDailyInputPayload" not in source
    assert "RestorePatchesFromRollback" not in source
    assert "Operational Apply/Restore executed: NO" in source


def test_native_probe_uses_macro_disabled_excel_for_all_non_vba_steps():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "_configure_excel(app)" in source
    assert "ReadOnly=True" in source
    assert "ReadOnly=False" in source
    assert "wb.Save" not in source
    assert "workbook.Save()" in source
    assert "DIOMIL_SYNTHETIC_FAULT" in source
    assert 'FICTIONAL_NATIVE_FAILURE' in source
    assert 'Phase.ROLLED_BACK, rollback_verified=True' in source


def test_native_probe_redacts_exception_detail(tmp_path, monkeypatch, capsys):
    source = tmp_path / "private-patient-name.xlsm"
    source.write_bytes(b"UNCHANGED")
    monkeypatch.setattr(probe.sys, "platform", "win32")
    def fail(_):
        raise RuntimeError("PRIVATE_PATIENT_DETAIL")
    monkeypatch.setattr(probe, "diagnose_native_rollback", fail)
    assert probe.main(["--source", str(source)]) == 2
    out = capsys.readouterr().out
    assert "SAFETY STOP" in out
    assert "PRIVATE_PATIENT_DETAIL" not in out
    assert "private-patient-name" not in out
    assert source.read_bytes() == b"UNCHANGED"
