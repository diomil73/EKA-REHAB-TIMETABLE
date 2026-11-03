from __future__ import annotations

PATIENT_REPORT_MODULE = "modPatientHistoryReport"
PATIENT_REPORT_SHEET = "PATIENT_HISTORY_REPORT"
PATIENT_REPORT_OPEN = "OpenPatientHistoryReport"
PATIENT_REPORT_NAV = "ekaToolbar_PatientHistory"
PATIENT_REPORT_PREFIX = "ekaPatientHistory_"

PATIENT_REPORT_VBA = r'''Option Explicit

Private Const REPORT_SHEET As String = "PATIENT_HISTORY_REPORT"

Public Sub OpenPatientHistoryReport()
    RenderPatientHistory "monthly"
End Sub

Public Sub ShowPatientHistoryDaily()
    RenderPatientHistory "daily"
End Sub

Public Sub ShowPatientHistoryWeekly()
    RenderPatientHistory "weekly"
End Sub

Public Sub ShowPatientHistoryMonthly()
    RenderPatientHistory "monthly"
End Sub

Public Sub ShowPatientHistoryOverall()
    RenderPatientHistory "overall"
End Sub

Private Sub RenderPatientHistory(ByVal reportMode As String)
    Dim scriptPath As String, outputPath As String, commandLine As String
    Dim ws As Worksheet, exitCode As Long, rawText As String
    Dim lines As Variant, fields As Variant, index As Long, rowIndex As Long, col As Long
    Dim priorEvents As Boolean, priorScreenUpdating As Boolean
    Dim succeeded As Boolean
    Dim previousSheet As Worksheet
    Dim failureNumber As Long, failureDescription As String

    priorEvents = Application.EnableEvents
    priorScreenUpdating = Application.ScreenUpdating
    On Error GoTo RenderError
    Set previousSheet = ThisWorkbook.ActiveSheet
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    scriptPath = ResolvePatientReportScript()
    If Len(scriptPath) = 0 Then Err.Raise vbObjectError + 7901, , "Patient report exporter not found"

    ThisWorkbook.Save

    outputPath = Environ$("TEMP") & "\\eka_patient_report_" & _
        Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000)) & ".tsv"
    commandLine = QuotePatientArg("python") & " " & QuotePatientArg(scriptPath) & _
        " --workbook " & QuotePatientArg(ThisWorkbook.FullName) & _
        " --output " & QuotePatientArg(outputPath) & _
        " --mode " & QuotePatientArg(reportMode)

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If exitCode <> 0 Or Dir$(outputPath) = "" Then _
        Err.Raise vbObjectError + 7902, , "Patient report export failed"

    rawText = Replace(ReadPatientUtf8(outputPath), ChrW(&HFEFF), "")
    lines = Split(Replace(rawText, vbCr, ""), vbLf)
    If UBound(lines) < 0 Then Err.Raise vbObjectError + 7903, , "Empty patient report"
    If CStr(lines(0)) <> "PeriodStart" & vbTab & "PeriodEnd" & vbTab & _
        "PatientId" & vbTab & "FinalizedDays" & vbTab & "Scheduled" & vbTab & _
        "Performed" & vbTab & "PatientAbsences" & vbTab & "Cancellations" & vbTab & _
        "Replacements" Then Err.Raise vbObjectError + 7904, , "Patient report header mismatch"

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Unprotect
    ' Previous report modes merge title/description rows. Unmerge before clearing,
    ' or Excel can raise 1004 while switching report modes.
    ws.Range("A1:I2").UnMerge
    ws.Range("A1:I1048576").ClearContents
    ws.Columns("C").NumberFormat = "@"

    ws.Range("A1").Value2 = "Ιστορικό / Στατιστικά Ασθενών"
    ws.Range("A2").Value2 = ModeDescription(reportMode)
    ws.Range("A5:I5").Value = Array( _
        "Από", "Έως", "Ασθενής", "Ημέρες", "Προγρ. συνεδρίες", _
        "Πραγμ. συνεδρίες", "Απουσίες", "Ακυρώσεις", "Αναπληρώσεις")

    rowIndex = 6
    For index = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(index)))) = 0 Then GoTo NextIndex
        fields = Split(CStr(lines(index)), vbTab)
        If UBound(fields) <> 8 Then Err.Raise vbObjectError + 7905, , "Invalid patient report row"
        For col = 0 To 8
            ws.Cells(rowIndex, col + 1).Value2 = fields(col)
        Next col
        If reportMode = "monthly" Then
            ws.Cells(rowIndex, 1).Value2 = Left$(fields(0), 7)
            ws.Cells(rowIndex, 2).Value2 = ""
        ElseIf reportMode = "daily" Then
            ws.Cells(rowIndex, 2).Value2 = ""
        End If
        rowIndex = rowIndex + 1
NextIndex:
    Next index

    FormatPatientReport ws, rowIndex - 1
    EnsurePatientButtons ws
    ws.Activate
    KeepApplicationShell
    ResetPatientReportWindow ws
    EnsurePatientButtons ws
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    succeeded = True

RenderCleanup:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    If Len(outputPath) > 0 And Dir$(outputPath) <> "" Then Kill outputPath
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorScreenUpdating
    On Error GoTo 0
    If Not succeeded Then
        On Error Resume Next
        If Not previousSheet Is Nothing Then previousSheet.Activate
        On Error GoTo 0
        MsgBox "Δεν ήταν δυνατή η προβολή του ιστορικού ασθενών. Δεν τροποποιήθηκε το ιστορικό." & _
            vbCrLf & "Κωδικός: " & CStr(failureNumber) & " / " & failureDescription, _
            vbCritical, "Ιστορικό ασθενών"
    End If
    Exit Sub

RenderError:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume RenderCleanup
End Sub

Private Function ModeDescription(ByVal mode As String) As String
    Select Case LCase$(Trim$(mode))
        Case "daily": ModeDescription = "Ημερήσια ανά κλεισμένη ημέρα"
        Case "weekly": ModeDescription = "Εβδομαδιαία, Δευτέρα έως Κυριακή"
        Case "monthly": ModeDescription = "Μηνιαία σύνοψη ανά ασθενή"
        Case "overall": ModeDescription = "Συνολικά από το διαθέσιμο ιστορικό"
        Case Else: ModeDescription = mode
    End Select
End Function

Private Sub EnsurePatientButtons(ByVal ws As Worksheet)
    EnsurePatientButton ws, "Daily", "Ημερήσια", "ShowPatientHistoryDaily", 98#
    EnsurePatientButton ws, "Weekly", "Εβδομαδιαία", "ShowPatientHistoryWeekly", 110#
    EnsurePatientButton ws, "Monthly", "Μηνιαία", "ShowPatientHistoryMonthly", 98#
    EnsurePatientButton ws, "Overall", "Συνολικά", "ShowPatientHistoryOverall", 90#
    EnsurePatientButton ws, "Back", "← MASTER", "GoToMaster", 86#
    PositionPatientButtons ws
End Sub

Private Sub EnsurePatientButton(ByVal ws As Worksheet, ByVal suffix As String, _
        ByVal caption As String, ByVal macro As String, ByVal width As Double)
    Dim shape As Shape
    Set shape = Nothing
    On Error Resume Next
    Set shape = ws.Shapes("ekaPatientHistory_" & suffix)
    On Error GoTo 0
    If shape Is Nothing Then
        Set shape = ws.Shapes.AddShape(5, 0, 0, width, 24)
        shape.Name = "ekaPatientHistory_" & suffix
    End If
    shape.OnAction = macro
    shape.Placement = 3
    shape.Visible = True
    shape.Width = width
    shape.Height = 24
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    On Error Resume Next
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shape.TextFrame2.TextRange.Font.Fill.Transparency = 0
    shape.ZOrder 0
    On Error GoTo 0
End Sub

Private Sub PositionPatientButtons(ByVal ws As Worksheet)
    Dim specs As Variant, i As Long, shape As Shape
    Dim leftPos As Double, topPos As Double
    specs = Array("Daily", "Weekly", "Monthly", "Overall", "Back")
    leftPos = ws.Range("A3").Left + 4
    topPos = ws.Rows(3).Top + 4
    For i = LBound(specs) To UBound(specs)
        Set shape = ws.Shapes("ekaPatientHistory_" & CStr(specs(i)))
        shape.Left = leftPos
        shape.Top = topPos
        shape.ZOrder 0
        leftPos = leftPos + shape.Width + 7
    Next i
End Sub

Private Sub FormatPatientReport(ByVal ws As Worksheet, ByVal lastRow As Long)
    On Error Resume Next
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Range("A1:I1").UnMerge
    ws.Range("A2:I2").UnMerge
    ws.Range("A1:I1").Merge
    ws.Range("A2:I2").Merge
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Font.Bold = True
    ws.Range("A5:I5").Font.Bold = True
    ws.Range("A5:I5").WrapText = True
    ws.Columns("A:B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 21
    ws.Columns("D:I").ColumnWidth = 17
    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 6
    ws.Rows(5).RowHeight = 30
    If lastRow >= 5 Then ws.Range("A5:I" & CStr(lastRow)).Borders.LineStyle = 1
    ws.ScrollArea = "A1:I" & CStr(Application.Max(lastRow + 2, 13))
    On Error GoTo 0
End Sub

Private Sub ResetPatientReportWindow(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitRow = 0
        ActiveWindow.SplitColumn = 0
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ws.Range("A6").Select
        ActiveWindow.SplitRow = 5
        ActiveWindow.FreezePanes = True
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Function ResolvePatientReportScript() As String
    Dim folder As String, candidate As String, i As Long
    folder = ThisWorkbook.Path
    For i = 0 To 2
        candidate = folder & "\\Python\\tools\\export_patient_history_ui.py"
        If Dir$(candidate) <> "" Then ResolvePatientReportScript = candidate: Exit Function
        folder = folder & "\\.."
    Next i
End Function

Private Function QuotePatientArg(ByVal value As String) As String
    QuotePatientArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function ReadPatientUtf8(ByVal fileName As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile fileName
    ReadPatientUtf8 = stream.ReadText
    stream.Close
End Function
'''


def install_patient_report_components(vbproject, workbook) -> None:
    """Install on a disposable native workbook, never an archived green proof."""
    try:
        old = vbproject.VBComponents(PATIENT_REPORT_MODULE)
        vbproject.VBComponents.Remove(old)
    except Exception:
        pass
    module = vbproject.VBComponents.Add(1)
    module.Name = PATIENT_REPORT_MODULE
    module.CodeModule.AddFromString(PATIENT_REPORT_VBA)

    try:
        report = workbook.Worksheets(PATIENT_REPORT_SHEET)
    except Exception:
        report = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        report.Name = PATIENT_REPORT_SHEET
    report.Unprotect()
    report.Cells.ClearContents()
    report.Visible = -1

    # Patient and therapist statistics belong together on THERAPIST_DAILY.
    # Fail closed if the companion H4 button was not installed in this workbook.
    source = workbook.Worksheets("THERAPIST_DAILY")
    therapist_button = source.Shapes("ekaToolbar_TherapistHistory")
    if str(therapist_button.OnAction).split("!")[-1].strip("'") != "OpenTherapistHistoryReport":
        raise RuntimeError("H4 therapist statistics button is not correctly bound")
    if not bool(therapist_button.Visible):
        raise RuntimeError("H4 therapist statistics button is not visible")

    # Remove legacy patient button from DAILY_INPUT if present.
    daily = workbook.Worksheets("DAILY_INPUT")
    try:
        daily.Shapes(PATIENT_REPORT_NAV).Delete()
    except Exception:
        pass
    try:
        source.Shapes(PATIENT_REPORT_NAV).Delete()
    except Exception:
        pass
    left = float(therapist_button.Left) + float(therapist_button.Width) + 8
    top = float(therapist_button.Top)
    shape = source.Shapes.AddShape(5, left, top, 146, 24)
    shape.Name = PATIENT_REPORT_NAV
    shape.OnAction = PATIENT_REPORT_OPEN
    shape.Placement = 3
    shape.Visible = True
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.TextFrame2.TextRange.Text = "Στατιστικά ασθενών"
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
