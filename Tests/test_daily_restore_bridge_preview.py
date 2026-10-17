from pathlib import Path
import inspect

import pytest

from Python.tools.build_daily_restore_bridge_preview import (
    DailyRestoreBridgePreviewError,
    _authoritative_ledger_rows,
    _build_parser,
    _sync_ledger_from_authoritative_facts,
    build_daily_restore_bridge_preview,
)


def test_bridge_preview_cli_requires_source_and_output():
    parser = _build_parser()
    args = parser.parse_args(["--source", "input.xlsm", "--output", "output.xlsm"])
    assert args.source == Path("input.xlsm")
    assert args.output == Path("output.xlsm")
    assert args.overwrite is False


def test_bridge_preview_refuses_missing_source(tmp_path):
    with pytest.raises(DailyRestoreBridgePreviewError, match="Source workbook not found"):
        build_daily_restore_bridge_preview(
            tmp_path / "missing.xlsm",
            tmp_path / "preview.xlsm",
        )


def test_bridge_preview_refuses_source_equals_output(tmp_path):
    source = tmp_path / "same.xlsm"
    source.write_bytes(b"not-an-xlsm")
    with pytest.raises(DailyRestoreBridgePreviewError, match="Output must be different"):
        build_daily_restore_bridge_preview(source, source)


def test_bridge_preview_uses_authoritative_facts_not_stale_ledger_count():
    source = inspect.getsource(_authoritative_ledger_rows)
    assert "read_daily_replacement_fact_rows(" in source
    assert "include_restored=True" in source
    assert "fact.status == ACTIVE_STATUS" in source
    assert "read_patient_registry(path)" in source
    assert "read_unified_base_schedule(path)" in source
    assert "materialize_sessions_for_date(base_entries, target_date)" in source


def test_ledger_sync_projects_fact_fields_and_preserves_audit_extras_by_id():
    source = inspect.getsource(_sync_ledger_from_authoritative_facts)
    assert "preserved[replacement_id]" in source
    assert "ledger.Cells(row_index, 12).Value = row[9]" in source
    assert "ledger.Cells(row_index, col).Value = value" in source
    assert "ClearContents()" in source
    assert "if offset == 12" in source


def test_bridge_preview_contract_syncs_ledger_and_installs_in_session_vba():
    source = inspect.getsource(build_daily_restore_bridge_preview)
    assert "shutil.copy2(source, output)" in source
    assert "_authoritative_ledger_rows(output)" in source
    assert "_sync_ledger_from_authoritative_facts(workbook, fact_rows)" in source
    assert "effective_vba = build_in_session_ledger_vba(VBA_CODE)" in source
    assert "_remove_component_if_present(vbproject, MODULE_NAME)" in source
    assert "module.CodeModule.AddFromString(effective_vba)" in source
    assert "RefreshDailyChangeLedgerView" in source
    assert "Source workbook changed during preview build" in source
    assert "output.unlink(missing_ok=True)" in source
