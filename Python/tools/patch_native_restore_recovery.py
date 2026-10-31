from __future__ import annotations

import argparse
from pathlib import Path
import sys


class NativeRestoreRecoveryPatchError(RuntimeError):
    pass


def _normalize_vba(code: str) -> str:
    return code.replace("\r\n", "\n").replace("\r", "\n")


def patch_daily_change_ledger_code(code: str) -> str:
    """Patch only the three native defects proven by the M smoke."""

    normalized = _normalize_vba(code)

    declaration_marker = (
        "    Dim rollbackPath As String, patchPath As String\n"
        "    Dim previousEvents As Boolean, previousScreenUpdating As Boolean"
    )
    declaration_replacement = (
        "    Dim rollbackPath As String, patchPath As String\n"
        "    Dim changeId As String\n"
        "    Dim previousEvents As Boolean, previousScreenUpdating As Boolean"
    )
    if declaration_marker not in normalized:
        raise NativeRestoreRecoveryPatchError(
            "daily-change restore declaration marker not found"
        )
    normalized = normalized.replace(declaration_marker, declaration_replacement, 1)

    apply_marker = (
        '        If Not ApplyPayloadFileInPlace(patchPath, "daily-change restore") Then '
        'Err.Raise vbObjectError + 2836, , "Restore patch apply failed"'
    )
    apply_replacement = (
        "        If Not ApplyValidatedDailyInputPayload(patchPath) Then "
        'Err.Raise vbObjectError + 2836, , "Restore patch apply failed"'
    )
    if apply_marker not in normalized:
        raise NativeRestoreRecoveryPatchError(
            "daily-change restore apply marker not found"
        )
    normalized = normalized.replace(apply_marker, apply_replacement, 1)

    overlap_marker = '''    startRow = wsDaily.Cells(wsDaily.Rows.Count, 1).End(xlUp).Row + 3
    If startRow < 21 Then startRow = 21
    wsDaily.Range("A" & startRow & ":G" & startRow + 200).ClearContents
    wsDaily.Range("A" & startRow & ":G" & startRow + 200).Font.Italic = False
    wsDaily.Range("A" & startRow & ":G" & startRow + 200).Font.Bold = False

    lastLedgerRow = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row
    If lastLedgerRow < 2 Then Exit Sub'''
    overlap_replacement = '''    lastLedgerRow = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row
    If lastLedgerRow < 2 Then Exit Sub

    ' Never clear an arbitrary block below column A. Operational detail can
    ' continue in other THERAPIST_DAILY columns. Place the ledger below the
    ' real last used row and only clear rows previously owned by the ledger.
    Dim previousStartRow As Long, previousEndRow As Long, previousDisplayRow As Long
    Dim lastDataCell As Range
    For ledgerRow = 2 To lastLedgerRow
        If IsNumeric(wsLedger.Cells(ledgerRow, 12).Value) Then
            previousDisplayRow = CLng(Val(CStr(wsLedger.Cells(ledgerRow, 12).Value)))
            If previousDisplayRow > 0 Then
                If previousStartRow = 0 Or previousDisplayRow < previousStartRow Then previousStartRow = previousDisplayRow
                If previousDisplayRow > previousEndRow Then previousEndRow = previousDisplayRow
            End If
        End If
    Next ledgerRow
    If previousStartRow > 0 Then
        previousStartRow = previousStartRow - 2
        If previousStartRow < 1 Then previousStartRow = 1
        previousEndRow = previousEndRow + 2
        wsDaily.Range("A" & previousStartRow & ":G" & previousEndRow).ClearContents
        wsDaily.Range("A" & previousStartRow & ":G" & previousEndRow).Font.Italic = False
        wsDaily.Range("A" & previousStartRow & ":G" & previousEndRow).Font.Bold = False
    End If

    Set lastDataCell = wsDaily.Cells.Find(What:="*", After:=wsDaily.Cells(1, 1), _
                                         LookIn:=xlFormulas, LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                         MatchCase:=False)
    If lastDataCell Is Nothing Then
        startRow = 21
    Else
        startRow = lastDataCell.Row + 3
        If startRow < 21 Then startRow = 21
    End If'''
    if overlap_marker not in normalized:
        raise NativeRestoreRecoveryPatchError(
            "THERAPIST_DAILY ledger overlap marker not found"
        )
    normalized = normalized.replace(overlap_marker, overlap_replacement, 1)

    if "ApplyPayloadFileInPlace" in normalized:
        raise NativeRestoreRecoveryPatchError(
            "non-canonical ApplyPayloadFileInPlace call remains"
        )
    if "Dim changeId As String" not in normalized:
        raise NativeRestoreRecoveryPatchError("restore changeId declaration missing")
    if 'wsDaily.Cells.Find(What:="*"' not in normalized:
        raise NativeRestoreRecoveryPatchError("safe ledger anchor was not installed")
    return normalized


def patch_daily_input_status_rich_text(code: str) -> str:
    """Clear character-level strike-through from the MASTER status/reason cell."""

    normalized = _normalize_vba(code)
    marker = "            master.Cells(row, statusCol).Font.Strikethrough = False"
    replacement = (
        "            master.Cells(row, statusCol).Font.Strikethrough = False\n"
        "            master.Cells(row, statusCol).Characters(1, "
        "Len(CStr(master.Cells(row, statusCol).Value))).Font.Strikethrough = False"
    )
    if normalized.count(marker) != 1:
        raise NativeRestoreRecoveryPatchError(
            "MASTER status normalization marker not found exactly once"
        )
    return normalized.replace(marker, replacement, 1)


def _replace_component_code(vbproject, component_name: str, code: str) -> None:
    component = vbproject.VBComponents(component_name)
    module = component.CodeModule
    count = int(module.CountOfLines)
    if count:
        module.DeleteLines(1, count)
    module.AddFromString(code)


def patch_workbook(workbook_path: Path) -> None:
    workbook_path = workbook_path.resolve()
    if not workbook_path.exists():
        raise NativeRestoreRecoveryPatchError(f"Workbook not found: {workbook_path}")
    if sys.platform != "win32":
        raise NativeRestoreRecoveryPatchError(
            "Native restore recovery patch requires Windows with Microsoft Excel"
        )

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise NativeRestoreRecoveryPatchError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(
            str(workbook_path), UpdateLinks=0, ReadOnly=False, AddToMru=False
        )
        vbproject = workbook.VBProject

        ledger = vbproject.VBComponents("modDailyChangeLedger").CodeModule
        ledger_code = str(ledger.Lines(1, int(ledger.CountOfLines)))
        patched_ledger = patch_daily_change_ledger_code(ledger_code)
        _replace_component_code(vbproject, "modDailyChangeLedger", patched_ledger)

        daily = vbproject.VBComponents("modDailyInputAction").CodeModule
        daily_code = str(daily.Lines(1, int(daily.CountOfLines)))
        patched_daily = patch_daily_input_status_rich_text(daily_code)
        _replace_component_code(vbproject, "modDailyInputAction", patched_daily)

        installed_ledger = _normalize_vba(
            str(
                vbproject.VBComponents("modDailyChangeLedger").CodeModule.Lines(
                    1,
                    int(vbproject.VBComponents("modDailyChangeLedger").CodeModule.CountOfLines),
                )
            )
        )
        installed_daily = _normalize_vba(
            str(
                vbproject.VBComponents("modDailyInputAction").CodeModule.Lines(
                    1,
                    int(vbproject.VBComponents("modDailyInputAction").CodeModule.CountOfLines),
                )
            )
        )
        if "ApplyPayloadFileInPlace" in installed_ledger:
            raise NativeRestoreRecoveryPatchError(
                "restore still references ApplyPayloadFileInPlace"
            )
        if "ApplyValidatedDailyInputPayload(patchPath)" not in installed_ledger:
            raise NativeRestoreRecoveryPatchError(
                "canonical restore patch apply call not installed"
            )
        if "Dim changeId As String" not in installed_ledger:
            raise NativeRestoreRecoveryPatchError("restore changeId declaration missing")
        if 'wsDaily.Cells.Find(What:="*"' not in installed_ledger:
            raise NativeRestoreRecoveryPatchError("safe ledger anchor missing")
        if ".Characters(1, Len(CStr(master.Cells(row, statusCol).Value))).Font.Strikethrough = False" not in installed_daily:
            raise NativeRestoreRecoveryPatchError(
                "MASTER status rich-text normalization missing"
            )

        workbook.Save()
    except NativeRestoreRecoveryPatchError:
        raise
    except Exception as exc:
        raise NativeRestoreRecoveryPatchError(
            f"Native restore recovery patch failed: {exc}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Patch the generated preview for native daily-change recovery defects."
    )
    parser.add_argument("--workbook", type=Path, required=True)
    args = parser.parse_args()
    try:
        patch_workbook(args.workbook)
    except NativeRestoreRecoveryPatchError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    print("NATIVE RESTORE RECOVERY PATCH OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
