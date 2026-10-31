from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys


class NativeRestoreRecoveryPatchError(RuntimeError):
    pass


def _normalize_vba(code: str) -> str:
    return code.replace("\r\n", "\n").replace("\r", "\n")


def _compact_vba(code: str) -> str:
    normalized = _normalize_vba(code)
    normalized = re.sub(r"_\s*\n", "", normalized)
    return re.sub(r"\s+", "", normalized)


_SAFE_LEDGER_REFRESH = r'''Public Sub RefreshDailyChangeLedgerView()
    Dim wsDaily As Worksheet, wsLedger As Worksheet
    Dim startRow As Long, lastLedgerRow As Long, ledgerRow As Long, displayRow As Long
    Dim previousStartRow As Long, previousEndRow As Long, previousDisplayRow As Long
    Dim lastDataCell As Range
    Dim shp As Shape, cb As Shape, btn As Shape
    Dim statusText As String, changeType As String

    On Error GoTo RenderError
    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)

    For Each shp In wsDaily.Shapes
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then shp.Delete
    Next shp

    lastLedgerRow = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row
    If lastLedgerRow < 2 Then Exit Sub

    ' Clear only the rows previously owned by this ledger UI.
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

    ' Anchor below the true last used row across the whole sheet, not column A.
    Set lastDataCell = wsDaily.Cells.Find(What:="*", After:=wsDaily.Cells(1, 1), _
                                         LookIn:=xlFormulas, LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                         MatchCase:=False)
    If lastDataCell Is Nothing Then
        startRow = 21
    Else
        startRow = lastDataCell.Row + 3
        If startRow < 21 Then startRow = 21
    End If

    wsDaily.Cells(startRow, 1).Value = "Αλλαγές ημέρας"
    With wsDaily.Range("A" & startRow & ":G" & startRow)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    wsDaily.Cells(startRow + 1, 1).Value = "✓"
    wsDaily.Cells(startRow + 1, 2).Value = "Ασθενής / Θεραπευτής"
    wsDaily.Cells(startRow + 1, 3).Value = "Από"
    wsDaily.Cells(startRow + 1, 4).Value = "Προς"
    wsDaily.Cells(startRow + 1, 5).Value = "Αιτία"
    wsDaily.Cells(startRow + 1, 6).Value = "Κατάσταση"
    wsDaily.Cells(startRow + 1, 7).Value = "Ενέργεια"
    wsDaily.Range("A" & startRow + 1 & ":G" & startRow + 1).Font.Bold = True

    displayRow = startRow + 2
    For ledgerRow = 2 To lastLedgerRow
        If Len(Trim$(CStr(wsLedger.Cells(ledgerRow, 1).Value))) = 0 Then GoTo NextLedger
        statusText = UCase$(Trim$(CStr(wsLedger.Cells(ledgerRow, 10).Value)))
        changeType = LCase$(Trim$(CStr(wsLedger.Cells(ledgerRow, 3).Value)))
        wsDaily.Cells(displayRow, 2).Value = wsLedger.Cells(ledgerRow, 5).Value
        wsDaily.Cells(displayRow, 3).Value = wsLedger.Cells(ledgerRow, 6).Value
        wsDaily.Cells(displayRow, 4).Value = wsLedger.Cells(ledgerRow, 7).Value
        wsDaily.Cells(displayRow, 5).Value = GreekLedgerReason(CStr(wsLedger.Cells(ledgerRow, 8).Value))
        If statusText = STATUS_RESTORED Then
            wsDaily.Cells(displayRow, 6).Value = "Επαναφέρθηκε"
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.Color = RGB(128, 128, 128)
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.Italic = True
        Else
            wsDaily.Cells(displayRow, 6).Value = "Ενεργή αλλαγή"
            If changeType = "patient_absence" Or changeType = "session_cancellation" Then
                Set cb = wsDaily.Shapes.AddFormControl(1, wsDaily.Cells(displayRow, 1).Left + 4, wsDaily.Cells(displayRow, 1).Top + 2, 16, 16)
                cb.Name = SHAPE_PREFIX & "Select_" & CStr(ledgerRow)
                cb.TextFrame.Characters.Text = ""
                Set btn = wsDaily.Shapes.AddFormControl(0, wsDaily.Cells(displayRow, 7).Left + 2, wsDaily.Cells(displayRow, 7).Top + 1, wsDaily.Cells(displayRow, 7).Width - 4, wsDaily.Cells(displayRow, 7).Height - 2)
                btn.Name = SHAPE_PREFIX & "Restore_" & CStr(ledgerRow)
                btn.TextFrame.Characters.Text = "Επαναφορά"
                btn.OnAction = "RestoreDailyChangeRow"
            Else
                wsDaily.Cells(displayRow, 7).Value = "Επαναφορά σύντομα"
            End If
        End If
        wsLedger.Cells(ledgerRow, 12).Value2 = displayRow
        displayRow = displayRow + 1
NextLedger:
    Next ledgerRow

    Set btn = wsDaily.Shapes.AddFormControl(0, wsDaily.Cells(displayRow + 1, 2).Left, wsDaily.Cells(displayRow + 1, 2).Top, 170, 26)
    btn.Name = SHAPE_PREFIX & "RestoreSelected"
    btn.TextFrame.Characters.Text = "Επαναφορά επιλεγμένων"
    btn.OnAction = "RestoreSelectedDailyChanges"

    Set btn = wsDaily.Shapes.AddFormControl(0, wsDaily.Cells(displayRow + 1, 5).Left, wsDaily.Cells(displayRow + 1, 5).Top, 145, 26)
    btn.Name = SHAPE_PREFIX & "RestoreAll"
    btn.TextFrame.Characters.Text = "Επαναφορά όλων"
    btn.OnAction = "RestoreAllDailyChanges"

    wsDaily.Columns("A").ColumnWidth = 8
    wsDaily.Columns("B").ColumnWidth = 28
    wsDaily.Columns("C:D").ColumnWidth = 24
    wsDaily.Columns("E").ColumnWidth = 30
    wsDaily.Columns("F").ColumnWidth = 18
    wsDaily.Columns("G").ColumnWidth = 21
    Exit Sub
RenderError:
    MsgBox "Δεν ήταν δυνατή η εμφάνιση της λίστας αλλαγών: " & Err.Description, vbExclamation, "Αλλαγές ημέρας"
End Sub'''


def _replace_vba_sub(code: str, signature: str, replacement: str) -> str:
    pattern = re.compile(
        rf"(?ims)^\s*{re.escape(signature)}\s*$.*?^\s*End Sub\s*$"
    )
    updated, count = pattern.subn(replacement, code, count=1)
    if count != 1:
        raise NativeRestoreRecoveryPatchError(
            f"VBA procedure boundary not found exactly once: {signature}"
        )
    return updated


def patch_daily_change_ledger_code(code: str) -> str:
    """Patch only the native defects proven by the M smoke."""

    normalized = _normalize_vba(code)

    normalized = _replace_vba_sub(
        normalized,
        "Public Sub RefreshDailyChangeLedgerView()",
        _SAFE_LEDGER_REFRESH,
    )

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

    apply_pattern = re.compile(
        r'(?m)^\s*If Not ApplyPayloadFileInPlace\(patchPath,\s*"daily-change restore"\) Then\s*Err\.Raise vbObjectError \+ 2836, , "Restore patch apply failed"\s*$'
    )
    apply_replacement = (
        '        If Not ApplyValidatedDailyInputPayload(patchPath) Then Err.Raise vbObjectError + 2836, , "Restore patch apply failed"'
    )
    normalized, apply_count = apply_pattern.subn(apply_replacement, normalized, count=1)
    if apply_count != 1:
        raise NativeRestoreRecoveryPatchError(
            "daily-change restore apply marker not found semantically"
        )

    if "ApplyPayloadFileInPlace" in normalized:
        raise NativeRestoreRecoveryPatchError(
            "non-canonical ApplyPayloadFileInPlace call remains"
        )
    if "Dim changeId As String" not in normalized:
        raise NativeRestoreRecoveryPatchError("restore changeId declaration missing")
    if "Dim lastDataCell As Range" not in normalized:
        raise NativeRestoreRecoveryPatchError("safe ledger declarations missing")
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
        if "Dim lastDataCell As Range" not in installed_ledger:
            raise NativeRestoreRecoveryPatchError("safe ledger declarations missing")
        if 'wsDaily.Cells.Find(What:="*"' not in installed_ledger:
            raise NativeRestoreRecoveryPatchError("safe ledger anchor missing")

        compact_daily = _compact_vba(installed_daily)
        rich_text_token = (
            "master.Cells(row,statusCol).Characters(1,Len(CStr(master.Cells(row,statusCol).Value)))"
            ".Font.Strikethrough=False"
        )
        if rich_text_token not in compact_daily:
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
