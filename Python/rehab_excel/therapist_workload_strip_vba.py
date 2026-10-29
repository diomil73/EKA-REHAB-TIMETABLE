from __future__ import annotations

MODULE_NAME = "modTherapistWorkloadStrip"
SHAPE_PREFIX = "ekaWorkload_"
BUTTON_NAME = SHAPE_PREFIX + "Refresh"  # legacy MASTER button name; removed by installer/runtime
NAV_MASTER_BUTTON_NAME = "ekaNav_Master"

MODULE_CODE = r'''Option Explicit

Private Const WORKLOAD_PREFIX As String = "ekaWorkload_"
Private Const DEFAULT_CAPACITY As Long = 6
Private Const NAV_MASTER_BUTTON As String = "ekaNav_Master"
Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const LEDGER_SHAPE_PREFIX As String = "ekaLedger_"

Public Sub ShowTherapistWorkload()
    On Error Resume Next
    ThisWorkbook.Worksheets("THERAPIST_DAILY").Activate
    On Error GoTo 0

    EnsureTherapistNavigation
    RenderTherapistWorkloadStrip
    KeepApplicationShell
End Sub

Public Sub GoToMasterFromTherapistDaily()
    On Error Resume Next
    ThisWorkbook.Worksheets("MASTER_SCHEDULE").Activate
    ApplyMasterView
    KeepApplicationShell
    On Error GoTo 0
End Sub

Public Sub EnsureTherapistNavigation()
    Dim daily As Worksheet
    Dim shape As Shape

    On Error GoTo NavigationExit
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")

    On Error Resume Next
    Set shape = daily.Shapes(NAV_MASTER_BUTTON)
    On Error GoTo NavigationExit

    If shape Is Nothing Then
        Set shape = daily.Shapes.AddShape(5, _
            daily.Range("A21").Left + 4, _
            daily.Range("A21").Top + 2, _
            170, 32)
        shape.Name = NAV_MASTER_BUTTON
    End If

    shape.OnAction = "GoToMasterFromTherapistDaily"
    shape.Placement = 3
    shape.Locked = True
    shape.TextFrame2.TextRange.Text = "Επιστροφή στο MASTER"
    StyleSupervisorButton shape, RGB(205, 235, 225), RGB(117, 166, 151)

    StyleTherapistSupervisorPanel

NavigationExit:
End Sub

Public Sub StyleTherapistSupervisorPanel()
    Dim daily As Worksheet
    Dim ledger As Worksheet
    Dim startRow As Long
    Dim lastLedgerRow As Long
    Dim ledgerRow As Long
    Dim lastDisplayRow As Long
    Dim displayRow As Long
    Dim actionRow As Long
    Dim rowIndex As Long
    Dim shp As Shape
    Dim shapeName As String
    Dim panelRange As Range

    On Error GoTo SafeExit
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    Set ledger = ThisWorkbook.Worksheets(LEDGER_SHEET)

    startRow = CLng(ledger.Range("R1").Value)
    If startRow < 2 Then Exit Sub

    lastLedgerRow = ledger.Cells(ledger.Rows.Count, 1).End(xlUp).Row
    lastDisplayRow = startRow + 2
    For ledgerRow = 2 To lastLedgerRow
        If IsNumeric(ledger.Cells(ledgerRow, 19).Value) Then
            displayRow = CLng(ledger.Cells(ledgerRow, 19).Value)
            If displayRow > lastDisplayRow Then lastDisplayRow = displayRow
        End If
    Next ledgerRow
    actionRow = lastDisplayRow + 2

    Set panelRange = daily.Range("A" & CStr(startRow) & ":H" & CStr(actionRow + 2))
    With panelRange
        .Interior.Color = RGB(242, 250, 247)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders.Color = RGB(222, 232, 228)
        .Borders.Weight = xlThin
    End With

    With daily.Range("A" & CStr(startRow) & ":H" & CStr(startRow))
        .Interior.Color = RGB(205, 235, 225)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(45, 87, 75)
        .RowHeight = 30
    End With
    daily.Cells(startRow, 1).Value = "Ημερήσιες αλλαγές και επαναφορές"

    With daily.Range("A" & CStr(startRow + 1) & ":G" & CStr(startRow + 1))
        .Interior.Color = RGB(226, 243, 237)
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(57, 91, 81)
        .RowHeight = 24
    End With

    For rowIndex = startRow + 2 To lastDisplayRow
        daily.Rows(rowIndex).RowHeight = 25
        If (rowIndex - startRow) Mod 2 = 0 Then
            daily.Range("A" & CStr(rowIndex) & ":G" & CStr(rowIndex)).Interior.Color = RGB(250, 253, 252)
        Else
            daily.Range("A" & CStr(rowIndex) & ":G" & CStr(rowIndex)).Interior.Color = RGB(238, 248, 244)
        End If
    Next rowIndex

    daily.Rows(actionRow).RowHeight = 38
    daily.Rows(actionRow + 1).RowHeight = 8

    On Error Resume Next
    Set shp = daily.Shapes(NAV_MASTER_BUTTON)
    On Error GoTo SafeExit
    If Not shp Is Nothing Then
        shp.Left = daily.Cells(actionRow, 1).Left + 2
        shp.Top = daily.Cells(actionRow, 1).Top + 2
        shp.Width = 170
        shp.Height = 32
        StyleSupervisorButton shp, RGB(205, 235, 225), RGB(117, 166, 151)
    End If

    For Each shp In daily.Shapes
        shapeName = CStr(shp.Name)
        If Left$(shapeName, Len(LEDGER_SHAPE_PREFIX & "Restore_")) = LEDGER_SHAPE_PREFIX & "Restore_" Then
            shp.Height = 24
            StyleSupervisorButton shp, RGB(249, 224, 224), RGB(204, 139, 139)
        ElseIf shapeName = LEDGER_SHAPE_PREFIX & "RestoreSelected" Then
            shp.Left = daily.Cells(actionRow, 3).Left + 2
            shp.Top = daily.Cells(actionRow, 3).Top + 2
            shp.Width = 190
            shp.Height = 32
            StyleSupervisorButton shp, RGB(249, 224, 224), RGB(204, 139, 139)
        ElseIf shapeName = LEDGER_SHAPE_PREFIX & "RestoreAll" Then
            shp.Left = daily.Cells(actionRow, 6).Left + 2
            shp.Top = daily.Cells(actionRow, 6).Top + 2
            shp.Width = 160
            shp.Height = 32
            StyleSupervisorButton shp, RGB(244, 207, 207), RGB(190, 113, 113)
        End If
    Next shp

    daily.Columns("A").ColumnWidth = 7
    daily.Columns("B").ColumnWidth = 25
    daily.Columns("C:D").ColumnWidth = 22
    daily.Columns("E").ColumnWidth = 27
    daily.Columns("F").ColumnWidth = 17
    daily.Columns("G").ColumnWidth = 20
    daily.Columns("H").ColumnWidth = 3

SafeExit:
End Sub

Private Sub StyleSupervisorButton( _
    ByVal shape As Shape, ByVal fillColor As Long, ByVal lineColor As Long)

    On Error Resume Next
    shape.Fill.ForeColor.RGB = fillColor
    shape.Fill.Solid
    shape.Line.ForeColor.RGB = lineColor
    shape.Line.Weight = 1.25
    shape.Locked = True
    shape.Placement = 3
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 10
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(70, 70, 70)
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
    On Error GoTo 0
End Sub

Public Sub RenderTherapistWorkloadStrip()
    Dim master As Worksheet
    Dim daily As Worksheet
    Dim headerRows As Variant
    Dim firstRows As Variant
    Dim lastRows As Variant
    Dim blockIndex As Long
    Dim col As Long
    Dim rowIndex As Long
    Dim providerName As String
    Dim providerKey As String
    Dim timeValue As String
    Dim occupied As Long
    Dim leftPos As Double
    Dim topPos As Double
    Dim badgeWidth As Double
    Dim gap As Double
    Dim badgeIndex As Long
    Dim shape As Shape
    Dim providers As Object
    Dim providerNames As Object
    Dim providerOrder As Collection
    Dim providerTimes As Object
    Dim orderIndex As Long

    On Error GoTo RenderError
    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")

    ' Privacy rule: MASTER must never expose therapist workload.
    DeleteWorkloadShapes master
    DeleteWorkloadShapes daily

    badgeWidth = 110
    gap = 5
    leftPos = daily.Range("K2").Left + 4
    topPos = daily.Range("K2").Top + 2

    headerRows = Array(1, 11)
    firstRows = Array(2, 12)
    lastRows = Array(8, 18)
    Set providers = CreateObject("Scripting.Dictionary")
    Set providerNames = CreateObject("Scripting.Dictionary")
    Set providerOrder = New Collection

    For blockIndex = LBound(headerRows) To UBound(headerRows)
        For col = 2 To 10
            providerName = Trim$(CStr(daily.Cells(headerRows(blockIndex), col).Value))
            If Len(providerName) > 0 Then
                providerKey = NormalizeProviderName(providerName)
                If Len(providerKey) > 0 Then
                    If Not providers.Exists(providerKey) Then
                        Set providerTimes = CreateObject("Scripting.Dictionary")
                        providers.Add providerKey, providerTimes
                        providerNames.Add providerKey, providerName
                        providerOrder.Add providerKey
                    Else
                        Set providerTimes = providers(providerKey)
                    End If

                    For rowIndex = firstRows(blockIndex) To lastRows(blockIndex)
                        If CellHasActiveWork(daily.Cells(rowIndex, col)) Then
                            timeValue = TimeKey(daily.Cells(rowIndex, 1).Value)
                            If Len(timeValue) = 0 Then timeValue = TimeKey(daily.Cells(rowIndex, 1).Text)
                            If Len(timeValue) > 0 Then providerTimes(timeValue) = True
                        End If
                    Next rowIndex
                End If
            End If
        Next col
    Next blockIndex

    badgeIndex = 0
    For orderIndex = 1 To providerOrder.Count
        providerKey = CStr(providerOrder(orderIndex))
        providerName = CStr(providerNames(providerKey))
        Set providerTimes = providers(providerKey)
        occupied = providerTimes.Count

        badgeIndex = badgeIndex + 1
        Set shape = daily.Shapes.AddShape(5, _
            leftPos, _
            topPos + (badgeIndex - 1) * (24 + gap), _
            badgeWidth, 24)
        shape.Name = WORKLOAD_PREFIX & "Badge" & CStr(badgeIndex)
        shape.Placement = 3
        shape.Locked = True
        shape.TextFrame2.TextRange.Text = providerName & " " & CStr(occupied) & "/" & CStr(DEFAULT_CAPACITY)
        shape.TextFrame2.TextRange.Font.Name = "Calibri"
        shape.TextFrame2.TextRange.Font.Size = 9
        shape.TextFrame2.TextRange.Font.Bold = True
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        shape.TextFrame2.VerticalAnchor = 3
        shape.Line.ForeColor.RGB = RGB(166, 166, 166)

        If occupied >= DEFAULT_CAPACITY Then
            shape.Fill.ForeColor.RGB = RGB(244, 176, 132)
        ElseIf occupied = DEFAULT_CAPACITY - 1 Then
            shape.Fill.ForeColor.RGB = RGB(255, 217, 102)
        Else
            shape.Fill.ForeColor.RGB = RGB(169, 208, 142)
        End If
    Next orderIndex

    EnsureTherapistNavigation
    Exit Sub

RenderError:
    MsgBox "Δεν ήταν δυνατή η εμφάνιση του φόρτου θεραπευτών: " & Err.Description, _
           vbExclamation, "Φόρτος θεραπευτών"
End Sub

Public Sub RemoveMasterTherapistWorkload()
    Dim master As Worksheet
    On Error Resume Next
    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    If Not master Is Nothing Then DeleteWorkloadShapes master
    On Error GoTo 0
End Sub

Private Function CellHasActiveWork(ByVal cell As Range) As Boolean
    Dim text As String
    Dim lineStart As Long
    Dim lineEnd As Long
    Dim firstChar As Long
    Dim ch As String
    Dim strikeValue As Variant
    Dim italicValue As Variant

    text = CStr(cell.Value)
    If Len(Trim$(text)) = 0 Then Exit Function

    lineStart = 1
    Do While lineStart <= Len(text)
        lineEnd = InStr(lineStart, text, vbLf)
        If lineEnd = 0 Then lineEnd = Len(text) + 1

        firstChar = lineStart
        Do While firstChar < lineEnd
            ch = Mid$(text, firstChar, 1)
            If ch <> " " And ch <> vbTab And ch <> vbCr Then Exit Do
            firstChar = firstChar + 1
        Loop

        If firstChar < lineEnd Then
            ch = Mid$(text, firstChar, 1)
            If ch <> "→" Then
                strikeValue = cell.Characters(firstChar, 1).Font.Strikethrough
                italicValue = cell.Characters(firstChar, 1).Font.Italic
                If Not IsNull(strikeValue) And Not IsNull(italicValue) Then
                    If strikeValue = False And italicValue = False Then
                        CellHasActiveWork = True
                        Exit Function
                    End If
                ElseIf IsNull(strikeValue) Or IsNull(italicValue) Then
                    If strikeValue <> True And italicValue <> True Then
                        CellHasActiveWork = True
                        Exit Function
                    End If
                End If
            End If
        End If

        lineStart = lineEnd + 1
    Loop
End Function

Private Function NormalizeProviderName(ByVal valueText As String) As String
    Dim text As String
    text = LCase$(Trim$(valueText))
    text = Replace(text, "ά", "α")
    text = Replace(text, "έ", "ε")
    text = Replace(text, "ή", "η")
    text = Replace(text, "ί", "ι")
    text = Replace(text, "ϊ", "ι")
    text = Replace(text, "ΐ", "ι")
    text = Replace(text, "ό", "ο")
    text = Replace(text, "ύ", "υ")
    text = Replace(text, "ϋ", "υ")
    text = Replace(text, "ΰ", "υ")
    text = Replace(text, "ώ", "ω")
    text = Replace(text, "ς", "σ")
    text = Replace(text, ".", "")
    text = Replace(text, " ", "")
    NormalizeProviderName = text
End Function

Private Function TimeKey(ByVal rawValue As Variant) As String
    Dim text As String
    On Error GoTo InvalidTime

    If IsDate(rawValue) Then
        TimeKey = Format$(CDate(rawValue), "hh:mm")
        Exit Function
    End If
    If IsNumeric(rawValue) Then
        TimeKey = Format$(CDate(CDbl(rawValue)), "hh:mm")
        Exit Function
    End If

    text = Replace(Trim$(CStr(rawValue)), ".", ":")
    If IsDate(text) Then TimeKey = Format$(TimeValue(text), "hh:mm")
    Exit Function

InvalidTime:
    TimeKey = ""
End Function

Private Sub DeleteWorkloadShapes(ByVal ws As Worksheet)
    Dim names As Collection
    Dim index As Long
    Dim shapeName As String
    Set names = New Collection

    For index = 1 To ws.Shapes.Count
        shapeName = CStr(ws.Shapes(index).Name)
        If Left$(shapeName, Len(WORKLOAD_PREFIX)) = WORKLOAD_PREFIX Then names.Add shapeName
    Next index

    For index = 1 To names.Count
        ws.Shapes(CStr(names(index))).Delete
    Next index
End Sub
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _style_button(shape, caption: str, macro: str) -> None:
    shape.OnAction = macro
    shape.Placement = 3
    shape.Locked = True
    shape.Fill.ForeColor.RGB = 14805965
    shape.Line.ForeColor.RGB = 9934453
    shape.Line.Weight = 1.25
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 10
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 4605510
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3


def _delete_workload_shapes(ws) -> None:
    names: list[str] = []
    for index in range(1, int(ws.Shapes.Count) + 1):
        shape = ws.Shapes(index)
        name = str(shape.Name)
        if name.startswith(SHAPE_PREFIX):
            names.append(name)
    for name in names:
        try:
            ws.Shapes(name).Delete()
        except Exception:
            pass


def install_therapist_workload_strip(vbproject, workbook) -> None:
    """Install supervisor workload view without exposing workload on MASTER."""

    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)

    master = workbook.Worksheets("MASTER_SCHEDULE")
    daily = workbook.Worksheets("THERAPIST_DAILY")

    # Privacy-safe migration: remove the old MASTER button and every old badge.
    try:
        master.Unprotect()
    except Exception:
        pass
    _delete_workload_shapes(master)

    try:
        daily.Shapes(NAV_MASTER_BUTTON_NAME).Delete()
    except Exception:
        pass
    nav = daily.Shapes.AddShape(
        5,
        float(daily.Range("A21").Left) + 4,
        float(daily.Range("A21").Top) + 2,
        170,
        32,
    )
    nav.Name = NAV_MASTER_BUTTON_NAME
    _style_button(nav, "Επιστροφή στο MASTER", "GoToMasterFromTherapistDaily")
