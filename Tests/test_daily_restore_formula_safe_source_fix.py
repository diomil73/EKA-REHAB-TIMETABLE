from __future__ import annotations

import importlib.util
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "Python" / "tools" / "apply_daily_restore_formula_safe_source_fix.py"


def _load_module():
    spec = importlib.util.spec_from_file_location("daily_restore_formula_safe_fix", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_restore_preparer_uses_presentation_only_for_master_non_active_overlay():
    module = _load_module()
    source = '''            if patch.source_tag == "master_daily_non_active_overlay":
                value = str(master[patch.cell].value or "")
                patches.append(
                    CellPatch(
                        sheet="MASTER_SCHEDULE",
                        cell=patch.cell,
                        value=value,
                        wrap_text=True,
                        text_runs=_plain_runs(value),
                        source_tag="master_daily_restore_overlay",
                    )
                )'''
    updated = module.patch_restore_preparer(source)
    assert "intent=WriteIntent.PRESENTATION" in updated
    assert "value=value" not in updated
    assert "text_runs=_plain_runs(value)" not in updated
    assert 'source_tag="master_daily_restore_overlay"' in updated


def test_daily_input_vba_adds_restore_presentation_contract_and_formula_safe_rollback():
    module = _load_module()
    source = '''        Case "master_daily_patient_status_overlay"
            cell.Font.Color = RGB(255, 0, 0)
    End Select

        Case "master_daily_patient_status_overlay"
            VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
        Case Else
            VerifySourceTagPresentation = True

    Dim rollbackBook As Workbook, parts As Variant
    Dim i As Long, sheetName As String, cellAddress As String
    Dim previousAlerts As Boolean

        rollbackBook.Worksheets(sheetName).Range(cellAddress).Copy Destination:=ThisWorkbook.Worksheets(sheetName).Range(cellAddress)
RestoreNext:'''
    updated = module.patch_daily_input_vba(source)
    assert 'Case "master_daily_restore_overlay", "master_daily_restore_status"' in updated
    assert "cell.Font.Strikethrough = False" in updated
    assert "CBool(strikeValue) = False" in updated
    assert "Dim rollbackCell As Range, targetCell As Range" in updated
    assert "rollbackCell.Copy Destination:=targetCell" in updated
    assert "If rollbackCell.HasFormula Then targetCell.Formula = rollbackCell.Formula" in updated


def test_canonical_sources_match_expected_post_fix_markers():
    module = _load_module()
    restore_text = module.RESTORE_PREPARER.read_text(encoding="utf-8")
    vba_text = module.DAILY_INPUT_VBA.read_text(encoding="utf-8")

    overlay_start = restore_text.index('if patch.source_tag == "master_daily_non_active_overlay":')
    overlay_end = restore_text.index('elif patch.source_tag == "master_daily_patient_status_overlay":')
    overlay_block = restore_text[overlay_start:overlay_end]

    assert 'intent=WriteIntent.PRESENTATION' in overlay_block
    assert 'source_tag="master_daily_restore_overlay"' in overlay_block
    assert 'value = str(master[patch.cell].value or "")' not in overlay_block
    assert 'text_runs=_plain_runs(value)' not in overlay_block

    status_start = overlay_end
    status_end = restore_text.index('finally:', status_start)
    status_block = restore_text[status_start:status_end]
    assert 'source_tag="master_daily_restore_status"' in status_block
    assert 'text_runs=_plain_runs(value)' in status_block

    assert 'Case "master_daily_restore_overlay", "master_daily_restore_status"' in vba_text
    assert "Dim rollbackCell As Range, targetCell As Range" in vba_text
    assert "rollbackCell.Copy Destination:=targetCell" in vba_text
    assert "If rollbackCell.HasFormula Then targetCell.Formula = rollbackCell.Formula" in vba_text
    assert (
        'rollbackBook.Worksheets(sheetName).Range(cellAddress).Copy Destination:=ThisWorkbook.Worksheets(sheetName).Range(cellAddress)'
        not in vba_text
    )
