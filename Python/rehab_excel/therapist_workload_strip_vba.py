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
    Dim occupied As Long
    Dim leftPos As Double
    Dim topPos As Double
    Dim badgeWidth As Double
    Dim buttonWidth As Double
    Dim gap As Double
    Dim badgeIndex As Long
    Dim shape As Shape

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
    badgeIndex = 0

    For blockIndex = LBound(headerRows) To UBound(headerRows)
        For col = 2 To 10
            providerName = Trim$(CStr(daily.Cells(headerRows(blockIndex), col).Value))
            If Len(providerName) > 0 Then
                occupied = 0
                For rowIndex = firstRows(blockIndex) To lastRows(blockIndex)
                    If Len(Trim$(CStr(daily.Cells(rowIndex, col).Value))) > 0 Then occupied = occupied + 1
                Next rowIndex

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
            End If
        Next col
    Next blockIndex

    Exit Sub

RenderError:
    MsgBox "Δεν ήταν δυνατή η εμφάνιση του φόρτου θεραπευτών: " & Err.Description, _
           vbExclamation, "Φόρτος θεραπευτών"
End Sub

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

    The strip counts occupied THERAPIST_DAILY timeslot cells, not patient lines,
    so legitimate two-patient sharing still consumes one operational timeslot.
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
