from __future__ import annotations

MODULE_NAME = "modTherapistWorkloadStrip"
SHAPE_PREFIX = "ekaWorkload_"
BUTTON_NAME = SHAPE_PREFIX + "Refresh"

MODULE_CODE = r'''Option Explicit

Private Const WORKLOAD_PREFIX As String = "ekaWorkload_"
Private Const DEFAULT_CAPACITY As Long = 6

Public Sub ShowTherapistWorkload()
    RenderTherapistWorkloadStrip
    On Error Resume Next
    ThisWorkbook.Worksheets("MASTER_SCHEDULE").Activate
    KeepApplicationShell
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
    Dim buttonWidth As Double
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

    DeleteWorkloadShapes master

    master.Rows(3).RowHeight = 28
    buttonWidth = 112
    badgeWidth = 92
    gap = 5
    leftPos = master.Range("A3").Left + 4
    topPos = master.Range("A3").Top + 2

    Set shape = master.Shapes.AddShape(5, leftPos, topPos, buttonWidth, 24)
    shape.Name = WORKLOAD_PREFIX & "Refresh"
    shape.OnAction = "ShowTherapistWorkload"
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = RGB(221, 235, 247)
    shape.Line.ForeColor.RGB = RGB(166, 166, 166)
    shape.TextFrame2.TextRange.Text = "Φόρτος θεραπευτών"
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3

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
                        If Len(Trim$(CStr(daily.Cells(rowIndex, col).Value))) > 0 Then
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
        Set shape = master.Shapes.AddShape(5, _
            leftPos + buttonWidth + gap + (badgeIndex - 1) * (badgeWidth + gap), _
            topPos, badgeWidth, 24)
        shape.Name = WORKLOAD_PREFIX & "Badge" & CStr(badgeIndex)
        shape.Placement = 3
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

    Exit Sub

RenderError:
    MsgBox "Δεν ήταν δυνατή η εμφάνιση του φόρτου θεραπευτών: " & Err.Description, _
           vbExclamation, "Φόρτος θεραπευτών"
End Sub

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


def install_therapist_workload_strip(vbproject, workbook) -> None:
    """Install the workload-strip macro and its MASTER refresh button.

    The strip aggregates each therapist by normalized name and counts unique
    occupied clock times across the current THERAPIST_DAILY layout blocks.
    Multiple patients sharing one therapist/time cell therefore still consume
    one operational timeslot, and a therapist represented in more than one
    layout block receives one combined badge rather than duplicate badges.
    """

    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)

    master = workbook.Worksheets("MASTER_SCHEDULE")
    master.Rows(3).RowHeight = 28

    try:
        master.Shapes(BUTTON_NAME).Delete()
    except Exception:
        pass

    left = float(master.Range("A3").Left) + 4
    top = float(master.Range("A3").Top) + 2
    shape = master.Shapes.AddShape(5, left, top, 112, 24)
    shape.Name = BUTTON_NAME
    shape.OnAction = "ShowTherapistWorkload"
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10921638
    shape.Line.Weight = 1
    shape.TextFrame2.TextRange.Text = "Φόρτος θεραπευτών"
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3