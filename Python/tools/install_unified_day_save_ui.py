from __future__ import annotations

import argparse
from pathlib import Path
import sys

MODULE_NAME = "modUnifiedDaySave"
MACRO_NAME = "SaveOperationalDay"
DAILY_BUTTON = "ekaToolbar_SaveOperationalDay"
MASTER_BUTTON = "ekaToolbar_SaveOperationalDayMaster"
OLD_BUTTONS = ("ekaToolbar_CloseDay", "ekaToolbar_DayAmendment")
SOURCE = r'''Option Explicit

' Route only. The two already-proven authoritative handlers remain unchanged.
Public Sub SaveOperationalDay()
    Dim ws As Worksheet, history As Worksheet
    Dim rawDate As Variant, serviceDate As Date
    Dim r As Long, lastRow As Long, matched As Long
    Dim status As String, recordedDate As Variant

    On Error GoTo SafetyStop
    If ThisWorkbook.ReadOnly Then GoTo SafetyStop

    Set ws = ThisWorkbook.Worksheets("DAILY_INPUT")
    rawDate = ws.Range("B2").Value
    If Not IsDate(rawDate) Then GoTo SafetyStop
    serviceDate = DateValue(CDate(rawDate))

    On Error Resume Next
    Set history = ThisWorkbook.Worksheets("FINALIZED_DAY_SNAPSHOTS")
    On Error GoTo SafetyStop

    If history Is Nothing Then
        CloseOperationalDay
        Exit Sub
    End If

    If CStr(history.Cells(1, 1).Value) <> "SnapshotId" Or _
       CStr(history.Cells(1, 2).Value) <> "ServiceDate" Or _
       CStr(history.Cells(1, 4).Value) <> "Status" Then GoTo SafetyStop

    lastRow = history.Cells(history.Rows.Count, 1).End(xlUp).Row
    matched = 0
    For r = 2 To lastRow
        recordedDate = history.Cells(r, 2).Value
        If IsDate(recordedDate) Then
            If DateValue(CDate(recordedDate)) = serviceDate Then
                status = UCase$(Trim$(CStr(history.Cells(r, 4).Value)))
                If status = "CURRENT_FINAL" Then matched = matched + 1
            End If
        ElseIf Trim$(CStr(recordedDate)) = Format$(serviceDate, "yyyy-mm-dd") Then
            status = UCase$(Trim$(CStr(history.Cells(r, 4).Value)))
            If status = "CURRENT_FINAL" Then matched = matched + 1
        End If
    Next r

    Select Case matched
        Case 0
            ' Existing history with this date but no current final is ambiguous.
            For r = 2 To lastRow
                recordedDate = history.Cells(r, 2).Value
                If Trim$(CStr(recordedDate)) = Format$(serviceDate, "yyyy-mm-dd") Then GoTo SafetyStop
                If IsDate(recordedDate) Then
                    If DateValue(CDate(recordedDate)) = serviceDate Then GoTo SafetyStop
                End If
            Next r
            CloseOperationalDay
        Case 1
            StartClosedDayAmendment
        Case Else
            GoTo SafetyStop
    End Select
    Exit Sub

SafetyStop:
    MsgBox "Δεν ήταν δυνατή η ασφαλής αναγνώριση της κατάστασης ημέρας." & vbCrLf & _
           "Δεν έγινε αποθήκευση. Επικοινωνήστε με τον διαχειριστή.", _
           vbCritical, "Αποθήκευση ημέρας"
End Sub
'''


class UnifiedDaySaveInstallError(RuntimeError):
    pass


def _delete_if_present(ws, name: str) -> None:
    try:
        ws.Shapes(name).Delete()
    except Exception:
        pass


def _add_shape(ws, name: str, caption: str, cell: str, width: int) -> None:
    anchor = ws.Range(cell)
    shape = ws.Shapes.AddShape(5, float(anchor.Left) + 2, float(anchor.Top) + 1, width, 25)
    shape.Name = name
    shape.OnAction = MACRO_NAME
    shape.Placement = 3
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 10
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3


def install(path: Path) -> None:
    if sys.platform != "win32":
        raise UnifiedDaySaveInstallError("Requires Windows Excel and pywin32")
    if not path.is_file() or path.suffix.lower() != ".xlsm":
        raise UnifiedDaySaveInstallError("Fresh .xlsm workbook required")
    import win32com.client  # type: ignore[import-not-found]

    excel = win32com.client.DispatchEx("Excel.Application")
    workbook = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        if workbook.ReadOnly:
            raise UnifiedDaySaveInstallError("Workbook opened read-only")
        components = workbook.VBProject.VBComponents
        for name in ("modCloseDayAction", "modDayAmendmentAction"):
            components(name)
        try:
            components.Remove(components(MODULE_NAME))
        except Exception:
            pass
        component = components.Add(1)
        component.Name = MODULE_NAME
        component.CodeModule.AddFromString(SOURCE)

        daily = workbook.Worksheets("DAILY_INPUT")
        master = workbook.Worksheets("MASTER_SCHEDULE")
        for name in (*OLD_BUTTONS, DAILY_BUTTON):
            _delete_if_present(daily, name)
        _delete_if_present(master, MASTER_BUTTON)
        _add_shape(daily, DAILY_BUTTON, "Αποθήκευση ημέρας", "A4", 154)
        # A single direct shortcut on MASTER after Apply & Continue returns there.
        _add_shape(master, MASTER_BUTTON, "Αποθήκευση ημέρας", "G2", 154)
        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()

    # Reopen persisted workbook to verify actual saved code and shape bindings.
    excel = win32com.client.DispatchEx("Excel.Application")
    workbook = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        code_module = workbook.VBProject.VBComponents(MODULE_NAME).CodeModule
        code = code_module.Lines(1, code_module.CountOfLines)
        if "Public Sub SaveOperationalDay()" not in code or "Case Else" not in code:
            raise UnifiedDaySaveInstallError("Persisted dispatcher source mismatch")
        daily = workbook.Worksheets("DAILY_INPUT")
        master = workbook.Worksheets("MASTER_SCHEDULE")
        for ws, name in ((daily, DAILY_BUTTON), (master, MASTER_BUTTON)):
            shape = ws.Shapes(name)
            if str(shape.OnAction).split("!")[-1].strip("'") != MACRO_NAME or not bool(shape.Visible):
                raise UnifiedDaySaveInstallError("Persisted button binding mismatch: " + name)
        for old_name in OLD_BUTTONS:
            try:
                daily.Shapes(old_name)
            except Exception:
                continue
            raise UnifiedDaySaveInstallError("Old duplicate button persists: " + old_name)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(description="Install a single state-aware Save Day control in a disposable workbook")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    try:
        install(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    print("UNIFIED DAY SAVE UI INSTALL OK")
    print("Controls: DAILY_INPUT + MASTER_SCHEDULE")
    print("First close: existing CloseOperationalDay")
    print("Amendment: existing StartClosedDayAmendment")
    print("Ambiguous state: safety stop")
    print("Persisted VBA reopen verification: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
