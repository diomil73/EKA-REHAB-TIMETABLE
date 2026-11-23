"""Mocked CLI wiring tests; real native Excel proof is Windows operator-only."""
from dataclasses import dataclass

from rehab_excel.authoritative_commit import file_sha256
from tools import verify_daily_input_full_rollback as cli


@dataclass(frozen=True)
class StubReport:
    snapshot_sha256: str
    restored_sha256: str
    checked_nonjournal_parts: int = 9
    original_journal_revisions: int = 1
    restored_journal_revisions: int = 3


def _files(tmp_path):
    a = tmp_path / "FICTIONAL_PREPARED.xlsm"
    b = tmp_path / "FICTIONAL_RESTORED.xlsm"
    a.write_bytes(b"synthetic snapshot")
    b.write_bytes(b"synthetic recovered")
    return a, b


def test_final_cli_calls_native_readonly_after_package_check(tmp_path, monkeypatch, capsys):
    a, b = _files(tmp_path)
    calls = []
    def package(*args, **kwargs):
        assert args == (a, b)
        assert kwargs == {"expected_snapshot_sha256": file_sha256(a)}
        calls.append("package")
        return StubReport(file_sha256(a), file_sha256(b))

    monkeypatch.setattr(cli, "audit_full_workbook_rollback", package)
    monkeypatch.setattr(cli, "_read_journal", lambda path: (("FAKE",),))
    def native(path, rows):
        assert path == b
        assert rows == (("FAKE",),)
        calls.append("native")
    monkeypatch.setattr(cli, "verify_native_excel_readonly", native)
    assert cli.main([
        "--snapshot", str(a), "--restored", str(b),
        "--snapshot-sha256", file_sha256(a), "--stage", "final",
    ]) == 0
    assert calls == ["package", "native"]
    output = capsys.readouterr().out
    assert "AUDIT GREEN" in output
    assert "DIAGNOSTIC ONLY" in output
    assert a.read_bytes() == b"synthetic snapshot"
    assert b.read_bytes() == b"synthetic recovered"


def test_pre_final_stage_is_never_mistaken_for_final_recovery(tmp_path, monkeypatch, capsys):
    a, b = _files(tmp_path)
    calls = []
    def before(*args, **kwargs):
        calls.append("before")
        return StubReport(file_sha256(a), file_sha256(b), restored_journal_revisions=2)
    monkeypatch.setattr(cli, "audit_restored_business_state_before_rollback", before)
    monkeypatch.setattr(cli, "audit_full_workbook_rollback", lambda *a, **k: 1 / 0)
    monkeypatch.setattr(cli, "_read_journal", lambda path: ())
    monkeypatch.setattr(cli, "verify_native_excel_readonly", lambda path, rows: calls.append("native"))
    assert cli.main([
        "--snapshot", str(a), "--restored", str(b),
        "--snapshot-sha256", file_sha256(a), "--stage", "before-final-row",
    ]) == 0
    assert calls == ["before", "native"]
    assert "Audit stage: before-final-row" in capsys.readouterr().out


def test_native_error_is_safety_stop_and_no_data_or_paths_leak(tmp_path, monkeypatch, capsys):
    a, b = _files(tmp_path)
    monkeypatch.setattr(
        cli, "audit_full_workbook_rollback",
        lambda *args, **kwargs: StubReport(file_sha256(a), file_sha256(b)),
    )
    monkeypatch.setattr(cli, "_read_journal", lambda path: ())
    def failed_native(*args):
        raise RuntimeError("PRIVATE_PATIENT_DETAIL")
    monkeypatch.setattr(cli, "verify_native_excel_readonly", failed_native)
    assert cli.main([
        "--snapshot", str(a), "--restored", str(b),
        "--snapshot-sha256", file_sha256(a), "--stage", "final",
    ]) == 2
    out = capsys.readouterr().out
    assert "SAFETY STOP" in out
    assert "PRIVATE_PATIENT_DETAIL" not in out
    assert str(a) not in out
    assert str(b) not in out


def test_cli_has_no_macro_enable_or_no_native_switch():
    from pathlib import Path
    source = Path(cli.__file__).read_text(encoding="utf-8")
    assert "verify_native_excel_readonly" in source
    assert "--skip-native" not in source
    assert "AutomationSecurity = 1" not in source
    assert "ThisWorkbook.Save" not in source
    assert "Operational Apply/Restore executed: NO" in source
