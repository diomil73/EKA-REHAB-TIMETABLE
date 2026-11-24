"""Synthetic contract tests for read-only Gate 2 inventory."""
from pathlib import Path

from rehab_excel.daily_input_journal_vba import JOURNAL_MODULE_CODE
from tools import diagnose_daily_input_guard_inventory_native as probe


def test_gate_source_exact_false():
    assert probe.inspect_guard_source(JOURNAL_MODULE_CODE)


def test_guard_rejects_true_or_additional_action():
    assert not probe.inspect_guard_source(
        JOURNAL_MODULE_CODE.replace(
            "DailyInputJournalApplyEnabled = False",
            "DailyInputJournalApplyEnabled = True",
        )
    )
    assert not probe.inspect_guard_source(
        JOURNAL_MODULE_CODE.replace(
            "DailyInputJournalApplyEnabled = False",
            "ThisWorkbook.Save\n    DailyInputJournalApplyEnabled = False",
        )
    )


def test_native_inventory_never_invokes_vba_or_operational_apply():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "TemporaryDirectory(" in source
    assert "create_probe_copy(source, work)" in source
    assert "_provision_copy(work" in source
    assert "ReadOnly=True" in source
    assert "file_sha256(source) != before" in source
    assert ".Run(" not in source
    assert "ApplyDailyInputPreview(" not in source
    assert "RestorePatchesFromRollback(" not in source


def test_nonwindows_stops_before_inspection(tmp_path, monkeypatch, capsys):
    filename = tmp_path / "fictional.xlsm"
    filename.write_bytes(b"FICTIONAL")
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(filename)]) == 2
    assert filename.read_bytes() == b"FICTIONAL"
    assert "SAFETY STOP" in capsys.readouterr().out
