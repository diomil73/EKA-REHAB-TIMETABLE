from __future__ import annotations

MODULE_NAME = "modSupervisorPanelStyle"

MODULE_CODE = r'''Option Explicit

Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const DAILY_SHEET As String = "THERAPIST_DAILY"
Private Const LEDGER_PREFIX As String = "ekaLedger_"
Private Const NAV_MASTER As String = "ekaNav_Master"
Private Const STATUS_ACTIVE As String = "ACTIVE"
Private Const DAILY_INPUT_PREFIX As String = "DAILY_INPUT:"

Public Sub ApplySupervisorPanelStyle()
    Dim ws As Worksheet
    Dim ledger As Worksheet
    Dim startRow As Long
    Dim lastLedgerRow As Long
    Dim ledgerRow As Long
    Dim displayRow As Long
    Dim lastDisplayRow As Long
    Dim actionRow As Long
    Dim statusText As String
    Dim metaText As String
    Dim panelRange As Range
    Dim rowRange As Range
    Dim btn As Shape

    On Error GoTo StyleExit
    Set ws = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set ledger = ThisWorkbook.Worksheets(LEDGER_SHEET)

    startRow = CLng(Val(CStr(ledger.Range("R1").Value)))
    If startRow < 2 Then Exit Sub

    lastLedgerRow = ledger.Cells(ledger.Rows.Count, 1).End(xlUp).Row
    lastDisplayRow = startRow + 1
    For ledgerRow = 2 To lastLedgerRow
        displayRow = CLng(Val(CStr(ledger.Cells(ledgerRow, 19).Value)))
        If displayRow > lastDisplayRow Then lastDisplayRow = displayRow
    Next ledgerRow
    actionRow = lastDisplayRow + 2

    Set panelRange = ws.Range("A" & CStr(startRow) & ":H" & CStr(actionRow))
    With panelRange
        .Interior.Color = RGB(244, 250, 248)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("A" & CStr(startRow) & ":H" & CStr(startRow))
        .Interior.Color = RGB(194, 229, 220)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(38, 86, 78)
    End With
    ws.Rows(startRow).RowHeight = 30

    With ws.Range("A" & CStr(startRow + 1) & ":G" & CStr(startRow + 1))
        .Interior.Color = RGB(222, 241, 235)
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(45, 78, 72)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(startRow + 1).RowHeight = 25

    For displayRow = startRow + 2 To lastDisplayRow
        Set rowRange = ws.Range("A" & CStr(displayRow) & ":G" & CStr(displayRow))
        If (displayRow - startRow) Mod 2 = 0 Then
            rowRange.Interior.Color = RGB(250, 253, 252)
        Else
            rowRange.Interior.Color = RGB(237, 247, 244)
        End If
        rowRange.Font.Size = 10
        ws.Rows(displayRow).RowHeight = 28
        With rowRange.Borders
            .Color = RGB(211, 226, 222)
            .Weight = xlHairline
        End With
    Next displayRow

    DeleteShapeIfPresent ws, LEDGER_PREFIX & "RestoreSelected"
    DeleteShapeIfPresent ws, LEDGER_PREFIX & "RestoreAll"
    DeleteShapeIfPresent ws, NAV_MASTER

    For ledgerRow = 2 To lastLedgerRow
        statusText = UCase$(Trim$(CStr(ledger.Cells(ledgerRow, 12).Value)))
        displayRow = CLng(Val(CStr(ledger.Cells(ledgerRow, 19).Value)))
        If statusText = STATUS_ACTIVE And displayRow >= startRow + 2 Then
            DeleteShapeIfPresent ws, LEDGER_PREFIX & "Restore_" & CStr(ledgerRow)
            Set btn = AddPanelButton( _
                ws, LEDGER_PREFIX & "Restore_" & CStr(ledgerRow), "Επαναφορά", _
                ws.Cells(displayRow, 7).Left + 6, ws.Cells(displayRow, 7).Top + 3, _
                Application.Max(96, ws.Cells(displayRow, 7).Width - 12), 22, False)
            metaText = Trim$(CStr(ledger.Cells(ledgerRow, 18).Value))
            If Left$(metaText, Len(DAILY_INPUT_PREFIX)) = DAILY_INPUT_PREFIX Then
                btn.OnAction = "RestoreDailyInputChangeRow"
            Else
                btn.OnAction = "RestoreDailyChangeRow"
            End If
        End If
    Next ledgerRow

    ws.Rows(actionRow).RowHeight = 42

    Set btn = AddPanelButton(ws, NAV_MASTER, "Επιστροφή στο MASTER", _
        ws.Cells(actionRow, 1).Left + 2, ws.Cells(actionRow, 1).Top + 4, 190, 32, True)
    btn.OnAction = "GoToMasterFromTherapistDaily"

    Set btn = AddPanelButton(ws, LEDGER_PREFIX & "RestoreSelected", "Επαναφορά επιλεγμένων", _
        ws.Cells(actionRow, 3).Left + 4, ws.Cells(actionRow, 3).Top + 4, 205, 32, False)
    btn.OnAction = "RestoreSelectedDailyChanges"

    Set btn = AddPanelButton(ws, LEDGER_PREFIX & "RestoreAll", "Επαναφορά όλων", _
        ws.Cells(actionRow, 5).Left + 4, ws.Cells(actionRow, 5).Top + 4, 175, 32, False)
    btn.OnAction = "RestoreAllDailyChanges"

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C:D").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 27
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 21

StyleExit:
End Sub

Private Function AddPanelButton( _
    ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
    ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
    ByVal heightValue As Double, ByVal mintButton As Boolean) As Shape

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, widthValue, heightValue)
    btn.Name = shapeName
    btn.Placement = 3
    btn.Locked = True
    btn.Line.Weight = 1.25

    If mintButton Then
        btn.Fill.ForeColor.RGB = RGB(204, 235, 226)
        btn.Line.ForeColor.RGB = RGB(102, 166, 151)
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(37, 88, 77)
    Else
        btn.Fill.ForeColor.RGB = RGB(248, 218, 214)
        btn.Line.ForeColor.RGB = RGB(207, 137, 130)
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(126, 55, 50)
    End If

    btn.TextFrame2.TextRange.Text = caption
    btn.TextFrame2.TextRange.Font.Name = "Calibri"
    btn.TextFrame2.TextRange.Font.Size = 10
    btn.TextFrame2.TextRange.Font.Bold = True
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    btn.TextFrame2.VerticalAnchor = 3
    Set AddPanelButton = btn
End Function

Private Sub DeleteShapeIfPresent(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def install_supervisor_panel_style(vbproject) -> None:
    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)
