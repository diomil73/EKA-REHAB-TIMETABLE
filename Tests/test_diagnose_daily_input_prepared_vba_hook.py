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



def test_vba_module_source_is_ascii_for_native_vbe_roundtrip():
    from rehab_excel.daily_input_journal_vba import JOURNAL_MODULE_CODE

    assert JOURNAL_MODULE_CODE.isascii()



def test_vbe_source_comparison_normalizes_only_line_endings_and_trailing_whitespace():
    crlf = chr(13) + chr(10)
    lf = chr(10)
    expected = crlf.join(["Option Explicit", "", "Public Sub Probe()  ", "End Sub"]) + crlf
    actual = lf.join(["Option Explicit", "", "Public Sub Probe()", "End Sub", "", ""])
    assert probe._canonical_vba_source(expected) == probe._canonical_vba_source(actual)
    assert "first_difference_line=none" in probe._module_source_mismatch(expected, actual)


def test_vbe_source_mismatch_reports_redacted_line_diagnostics():
    lf = chr(10)
    detail = probe._module_source_mismatch(
        lf.join(["Option Explicit", "Public Sub A()", "End Sub"]),
        lf.join(["Option Explicit", "Public Sub B()", "End Sub"]),
    )
    assert "first_difference_line=2" in detail
    assert "expected_line_length=14" in detail
    assert "actual_line_length=14" in detail
    assert "differing_characters=1" in detail
    assert "character_diff_samples=12:U+0041>U+0042" in detail
    assert "sha256=" in detail
    assert "Public Sub A" not in detail
