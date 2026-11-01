from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
RESTORE_PREPARER = REPO_ROOT / "Python" / "tools" / "prepare_daily_change_restore_in_place.py"
DAILY_INPUT_VBA = REPO_ROOT / "Python" / "rehab_excel" / "daily_input_action_vba.py"


class RestoreFormulaSafeSourceFixError(RuntimeError):
    pass


def _replace_once(text: str, old: str, new: str, *, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise RestoreFormulaSafeSourceFixError(
            f"{label}: expected source marker exactly once, found {count}"
        )
    return text.replace(old, new, 1)


def patch_restore_preparer(text: str) -> str:
    old = '''            if patch.source_tag == "master_daily_non_active_overlay":
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
    new = '''            if patch.source_tag == "master_daily_non_active_overlay":
                # The daily non-active overlay never replaces the underlying MASTER
                # treatment formula; it only changes presentation.  Restore it the
                # same way.  Writing the formula text as a value makes VBA compare
                # cell.Value (calculated result) with the formula string and causes
                # a deterministic read-back mismatch.
                patches.append(
                    CellPatch(
                        sheet="MASTER_SCHEDULE",
                        cell=patch.cell,
                        intent=WriteIntent.PRESENTATION,
                        wrap_text=True,
                        source_tag="master_daily_restore_overlay",
                    )
                )'''
    return _replace_once(
        text,
        old,
        new,
        label="clean-day MASTER restore contract",
    )


def patch_daily_input_vba(text: str) -> str:
    apply_old = '''        Case "master_daily_patient_status_overlay"\n            cell.Font.Color = RGB(255, 0, 0)\n    End Select'''
    apply_new = '''        Case "master_daily_patient_status_overlay"\n            cell.Font.Color = RGB(255, 0, 0)\n        Case "master_daily_restore_overlay", "master_daily_restore_status"\n            cell.Font.Color = RGB(0, 0, 0)\n            cell.Font.Strikethrough = False\n    End Select'''
    text = _replace_once(
        text,
        apply_old,
        apply_new,
        label="MASTER restore presentation apply",
    )

    verify_old = '''        Case "master_daily_patient_status_overlay"\n            VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))\n        Case Else\n            VerifySourceTagPresentation = True'''
    verify_new = '''        Case "master_daily_patient_status_overlay"\n            VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))\n        Case "master_daily_restore_overlay", "master_daily_restore_status"\n            strikeValue = cell.Font.Strikethrough\n            If IsNull(strikeValue) Then Exit Function\n            VerifySourceTagPresentation = _\n                (CBool(strikeValue) = False And CLng(cell.Font.Color) = CLng(RGB(0, 0, 0)))\n        Case Else\n            VerifySourceTagPresentation = True'''
    text = _replace_once(
        text,
        verify_old,
        verify_new,
        label="MASTER restore presentation verify",
    )

    dim_old = '''    Dim rollbackBook As Workbook, parts As Variant\n    Dim i As Long, sheetName As String, cellAddress As String\n    Dim previousAlerts As Boolean'''
    dim_new = '''    Dim rollbackBook As Workbook, parts As Variant\n    Dim rollbackCell As Range, targetCell As Range\n    Dim i As Long, sheetName As String, cellAddress As String\n    Dim previousAlerts As Boolean'''
    text = _replace_once(text, dim_old, dim_new, label="rollback range declarations")

    copy_old = '''        rollbackBook.Worksheets(sheetName).Range(cellAddress).Copy Destination:=ThisWorkbook.Worksheets(sheetName).Range(cellAddress)\nRestoreNext:'''
    copy_new = '''        Set rollbackCell = rollbackBook.Worksheets(sheetName).Range(cellAddress)\n        Set targetCell = ThisWorkbook.Worksheets(sheetName).Range(cellAddress)\n        rollbackCell.Copy Destination:=targetCell\n        ' Cross-workbook Copy can rewrite internal formulas as [1]Sheet! references.\n        ' Re-assign the original formula text after copying formats/content.\n        If rollbackCell.HasFormula Then targetCell.Formula = rollbackCell.Formula\nRestoreNext:'''
    text = _replace_once(text, copy_old, copy_new, label="formula-safe rollback copy")
    return text


def _apply(path: Path, transform, *, label: str) -> bool:
    text = path.read_text(encoding="utf-8")
    updated = transform(text)
    if updated == text:
        raise RestoreFormulaSafeSourceFixError(f"{label}: no source change produced")
    path.write_text(updated, encoding="utf-8")
    print(f"[OK] {label}: {path.relative_to(REPO_ROOT)}")
    return True


def main() -> int:
    try:
        _apply(RESTORE_PREPARER, patch_restore_preparer, label="restore preparer")
        _apply(DAILY_INPUT_VBA, patch_daily_input_vba, label="DAILY_INPUT VBA")
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DAILY RESTORE FORMULA-SAFE SOURCE FIX APPLIED")
    print("NEXT: run focused regression tests before committing canonical source.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
