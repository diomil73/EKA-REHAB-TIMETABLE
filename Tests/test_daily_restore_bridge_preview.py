from pathlib import Path
import inspect

import pytest

from Python.tools.build_daily_restore_bridge_preview import (
    DailyRestoreBridgePreviewError,
    _build_parser,
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


def test_bridge_preview_contract_preserves_ledger_and_refreshes_only_vba_module():
    source = inspect.getsource(build_daily_restore_bridge_preview)
    assert "shutil.copy2(source, output)" in source
    assert "workbook.Worksheets(LEDGER_SHEET)" in source
    assert 'status == "ACTIVE"' in source
    assert "_remove_component_if_present(vbproject, MODULE_NAME)" in source
    assert "module.CodeModule.AddFromString(VBA_CODE)" in source
    assert "RefreshDailyChangeLedgerView" in source
    assert "Source workbook changed during preview build" in source
    assert "output.unlink(missing_ok=True)" in source
