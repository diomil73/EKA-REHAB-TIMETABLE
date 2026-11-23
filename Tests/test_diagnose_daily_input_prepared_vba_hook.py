from pathlib import Path

from rehab_excel.daily_input_operation import DailyInputOperationPhase as Phase
from tools import diagnose_daily_input_prepared_vba_hook as probe


def test_vba_probe_history_contains_two_fictional_terminal_chains():
    states = list(probe._fictional_operations("2026-11-23"))
    assert len(states) == 6
    assert [state.phase for state in states] == [
        Phase.PREPARED, Phase.FAILED, Phase.ROLLED_BACK,
        Phase.PREPARED, Phase.FAILED, Phase.ROLLED_BACK,
    ]
    assert all(state.rollback_verified for state in states if state.phase == Phase.ROLLED_BACK)
    assert all("FICTIONAL" in state.operation_id for state in states)


def test_probe_tool_refuses_non_windows_without_touching_source(tmp_path, monkeypatch, capsys):
    source = tmp_path / "synthetic.xlsm"
    source.write_bytes(b"synthetic-only")
    before = source.read_bytes()
    monkeypatch.setattr(probe.sys, "platform", "linux")
    assert probe.main(["--source", str(source)]) == 2
    assert "requires Windows + Excel" in capsys.readouterr().out
    assert source.read_bytes() == before


def test_probe_tool_uses_macro_disabled_open_and_only_explicit_named_probe():
    source = Path(probe.__file__).read_text(encoding="utf-8")
    assert "app.AutomationSecurity = 3" in source
    assert "app.AutomationSecurity = 1" in source
    assert "_assert_no_auto_open_macros(workbook)" in source
    assert "Auto_Open" in source
    assert "RunPreparedDailyInputJournalProbe" in source
    assert "DailyInputJournalApplyEnabled" in source
    assert "Operational Apply/Restore executed: NO" in source
    assert "ApplyDailyInputPreview" not in source
