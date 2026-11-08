from datetime import date

from tools import audit_h10_baseline_provenance as mod


def test_provenance_audit_identical_original_day_is_read_only(tmp_path, monkeypatch, capsys):
    original = tmp_path / "AZ.xlsm"
    candidate = tmp_path / "CA.xlsm"
    original.write_bytes(b"original")
    candidate.write_bytes(b"candidate")
    summary = (("SNAP", 1, "THERAPIST", 6, 7, 7, 7, 7, 0, 7, 7, ("08:00",)),)
    monkeypatch.setattr(mod, "_summaries", lambda path, day: summary)
    assert mod.audit(original, candidate, date(2026, 10, 30)) == 0
    assert b"original" == original.read_bytes()
    assert b"candidate" == candidate.read_bytes()
    assert "BASELINE DAY IDENTICAL" in capsys.readouterr().out


def test_provenance_audit_mismatched_history_fails_closed(tmp_path, monkeypatch, capsys):
    original = tmp_path / "AZ.xlsm"
    candidate = tmp_path / "CA.xlsm"
    original.write_bytes(b"original")
    candidate.write_bytes(b"candidate")
    monkeypatch.setattr(
        mod, "_summaries",
        lambda path, day: ((1, 2, 3),) if path == original else ((1, 2, 4),),
    )
    assert mod.audit(original, candidate, date(2026, 10, 30)) == 3
    assert "SAFETY STOP" in capsys.readouterr().out


def test_provenance_rejects_missing_original(tmp_path):
    candidate = tmp_path / "CA.xlsm"
    candidate.write_bytes(b"candidate")
    assert mod.audit(tmp_path / "missing.xlsm", candidate, date(2026, 10, 30)) == 2
