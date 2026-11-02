from __future__ import annotations

THERAPIST_HISTORY_MODULE_NAME = "modTherapistHistoryReport"
THERAPIST_HISTORY_SHEET_NAME = "THERAPIST_HISTORY_REPORT"
THERAPIST_HISTORY_OPEN_MACRO = "OpenTherapistHistoryReport"
THERAPIST_HISTORY_NAV_BUTTON = "ekaToolbar_TherapistHistory"
THERAPIST_HISTORY_BUTTON_PREFIX = "ekaHistory_"

THERAPIST_HISTORY_MODULE_CODE = r'''Option Explicit

Private Const REPORT_SHEET As String = "THERAPIST_HISTORY_REPORT"

Public Sub OpenTherapistHistoryReport()
    RenderTherapistHistoryReport "monthly"
End Sub

Public Sub ShowTherapistHistoryDaily()
    RenderTherapistHistoryReport "daily"
End Sub

Public Sub ShowTherapistHistoryWeekly()
    RenderTherapistHistoryReport "weekly"
End Sub

Public Sub ShowTherapistHistoryMonthly()
    RenderTherapistHistoryReport "monthly"
End Sub

Public Sub ShowTherapistHistoryOverall()
    RenderTherapistHistoryReport "overall"
End Sub

Private Sub RenderTherapistHistoryReport(ByVal reportMode As String)
    Dim exportScript As String, outputPath As String, commandLine As String
    Dim exitCode As Long, rawText As String, lines As Variant, parts As Variant
    Dim ws As Worksheet, rowIndex As Long, sourceLine As Long, col As Long
    Dim headers As Variant

    On Error GoTo ReportError

    exportScript = ResolveReportTool("export_therapist_history_ui.py")
    If Len(exportScript) = 0 Then
        MsgBox "Η προβολή ιστορικού δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Ιστορικό θεραπευτών"
        Exit Sub
    End If

    Set ws = ExistingReportSheet()
    If ws Is Nothing Then
        MsgBox "Δεν βρέθηκε η οθόνη ιστορικού θεραπευτών.", vbCritical, "Ιστορικό θεραπευτών"
        Exit Sub
    End If

    ThisWorkbook.Save
    outputPath = Environ$("TEMP") & "\\eka_therapist_history_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".tsv"
    commandLine = QuoteReportArg("python") & " " & QuoteReportArg(exportScript) & _
                  " --workbook " & QuoteReportArg(ThisWorkbook.FullName) & _
                  " --output " & QuoteReportArg(outputPath) & _
                  " --mode " & QuoteReportArg(reportMode)

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If exitCode <> 0 Or Dir$(outputPath) = "" Then
        MsgBox "Δεν ήταν δυνατή η δημιουργία του ιστορικού θεραπευτών.", vbExclamation, "Ιστορικό θεραπευτών"
        GoTo CleanUp
    End If

    rawText = Replace(ReadReportUtf8File(outputPath), ChrW(&HFEFF), "")
    lines = Split(Replace(rawText, vbCr, ""), vbLf)
    If UBound(lines) < 0 Or Len(Trim$(CStr(lines(0)))) = 0 Then
        MsgBox "Το ιστορικό θεραπευτών δεν επέστρεψε δεδομένα.", vbExclamation, "Ιστορικό θεραπευτών"
        GoTo CleanUp
    End If

    headers = ReportHeaders(reportMode)

    ws.Unprotect
    On Error Resume Next
    ws.Range("A1:N2").UnMerge
    On Error GoTo ReportError
    ws.Range("A1:M1048576").ClearContents
    EnsureReportButtons ws
    ws.Cells(1, 1).Value2 = "Ιστορικό / Στατιστικά Θεραπευτών"
    ws.Cells(2, 1).Value2 = ReportModeDescription(reportMode)
    For col = 0 To UBound(headers)
        ws.Cells(5, col + 1).Value2 = headers(col)
    Next col

    rowIndex = 6
    For sourceLine = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(sourceLine)))) = 0 Then GoTo NextLine
        parts = Split(CStr(lines(sourceLine)), vbTab)
        If UBound(parts) <> 13 Then Err.Raise vbObjectError + 7601, , "Invalid therapist-history UI row"
        ws.Cells(rowIndex, 1).Value2 = DisplayPeriodStart(reportMode, parts(0), parts(1))
        ws.Cells(rowIndex, 2).Value2 = DisplayPeriodEnd(reportMode, parts(0), parts(1))
        ws.Cells(rowIndex, 3).Value2 = parts(3)
        ws.Cells(rowIndex, 4).Value2 = parts(4)
        For col = 5 To 12
            ws.Cells(rowIndex, col).Value2 = parts(col)
        Next col
        If Len(parts(12)) > 0 Then ws.Cells(rowIndex, 12).Value2 = parts(12) & "%"
        If Len(parts(13)) > 0 Then ws.Cells(rowIndex, 13).Value2 = parts(13) & "%"
        rowIndex = rowIndex + 1
NextLine:
    Next sourceLine

    FormatReportSheet ws, rowIndex - 1
    EnsureReportButtons ws
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    ws.Activate
    ws.Range("A1").Select
    KeepApplicationShell

CleanUp:
    On Error Resume Next
    If Len(outputPath) > 0 And Dir$(outputPath) <> "" Then Kill outputPath
    On Error GoTo 0
    Exit Sub

ReportError:
    MsgBox "Δεν ήταν δυνατή η ενημέρωση του ιστορικού θεραπευτών.", vbCritical, "Ιστορικό θεραπευτών"
    Resume CleanUp
End Sub

Private Function ReportHeaders(ByVal reportMode As String) As Variant
    Select Case LCase$(Trim$(reportMode))
        Case "daily"
            ReportHeaders = Array("Ημερομηνία", "", "Θεραπευτής", "Ημέρες", "Προγρ. θέσεις", "Πραγμ. θέσεις", "Προγρ. συνεδρίες", "Πραγμ. συνεδρίες", "Αναπληρώσεις", "Αποδεσμεύσεις", "Αναστολές", "Προγρ. φόρτος", "Πραγμ. φόρτος")
        Case "monthly"
            ReportHeaders = Array("Μήνας", "", "Θεραπευτής", "Ημέρες", "Προγρ. θέσεις", "Πραγμ. θέσεις", "Προγρ. συνεδρίες", "Πραγμ. συνεδρίες", "Αναπληρώσεις", "Αποδεσμεύσεις", "Αναστολές", "Προγρ. φόρτος", "Πραγμ. φόρτος")
        Case Else
            ReportHeaders = Array("Από", "Έως", "Θεραπευτής", "Ημέρες", "Προγρ. θέσεις", "Πραγμ. θέσεις", "Προγρ. συνεδρίες", "Πραγμ. συνεδρίες", "Αναπληρώσεις", "Αποδεσμεύσεις", "Αναστολές", "Προγρ. φόρτος", "Πραγμ. φόρτος")
    End Select
End Function

Private Function ReportModeDescription(ByVal reportMode As String) As String
    Select Case LCase$(Trim$(reportMode))
        Case "daily": ReportModeDescription = "Ημερήσια προβολή ανά κλεισμένη ημέρα"
        Case "weekly": ReportModeDescription = "Εβδομαδιαία σύνοψη (Δευτέρα έως Κυριακή)"
        Case "monthly": ReportModeDescription = "Μηνιαία σύνοψη"
        Case "overall": ReportModeDescription = "Συνολικά για όλο το διαθέσιμο ιστορικό"
        Case Else: ReportModeDescription = reportMode
    End Select
End Function

Private Function DisplayPeriodStart(ByVal reportMode As String, ByVal startText As String, ByVal endText As String) As String
    Select Case LCase$(Trim$(reportMode))
        Case "monthly": DisplayPeriodStart = Left$(startText, 7)
        Case Else: DisplayPeriodStart = startText
    End Select
End Function

Private Function DisplayPeriodEnd(ByVal reportMode As String, ByVal startText As String, ByVal endText As String) As String
    Select Case LCase$(Trim$(reportMode))
        Case "daily", "monthly": DisplayPeriodEnd = ""
        Case Else: DisplayPeriodEnd = endText
    End Select
End Function

Private Sub FormatReportSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    On Error Resume Next
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Range("A1:M1").Merge
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:M2").Merge
    ws.Range("A2").Font.Bold = True
    ws.Range("A5:M5").Font.Bold = True
    ws.Range("A5:M5").WrapText = True
    ws.Columns("A:B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 19
    ws.Columns("D:M").ColumnWidth = 13
    If lastRow >= 5 Then ws.Range("A5:M" & CStr(lastRow)).Borders.LineStyle = 1
    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 6
    ws.Rows(5).RowHeight = 30
    PositionReportButtons ws
    ws.ScrollArea = "A1:M" & CStr(Application.Max(lastRow + 2, 13))
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.FreezePanes = False
        ws.Range("A6").Select
        ActiveWindow.SplitRow = 5
        ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
End Sub

Private Sub EnsureReportButtons(ByVal ws As Worksheet)
    EnsureOneReportButton ws, "ekaHistory_Daily", "Ημερήσια", "ShowTherapistHistoryDaily", 98#
    EnsureOneReportButton ws, "ekaHistory_Weekly", "Εβδομαδιαία", "ShowTherapistHistoryWeekly", 108#
    EnsureOneReportButton ws, "ekaHistory_Monthly", "Μηνιαία", "ShowTherapistHistoryMonthly", 98#
    EnsureOneReportButton ws, "ekaHistory_Overall", "Συνολικά", "ShowTherapistHistoryOverall", 90#
    EnsureOneReportButton ws, "ekaHistory_Back", "← MASTER", "GoToMaster", 86#
    PositionReportButtons ws
End Sub

Private Sub EnsureOneReportButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal caption As String, ByVal macroName As String, ByVal buttonWidth As Double)
    Dim shape As Shape
    On Error Resume Next
    Set shape = ws.Shapes(buttonName)
    On Error GoTo 0
    If shape Is Nothing Then
        Set shape = ws.Shapes.AddShape(5, 0, 0, buttonWidth, 24)
        shape.Name = buttonName
    End If
    shape.OnAction = macroName
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.Line.Weight = 1
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    On Error Resume Next
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
    shape.ZOrder 0
    On Error GoTo 0
End Sub

Private Sub PositionReportButtons(ByVal ws As Worksheet)
    Dim specs As Variant, i As Long
    Dim shape As Shape, leftPos As Double, topPos As Double, buttonWidth As Double
    specs = Array( _
        Array("ekaHistory_Daily", 98#), _
        Array("ekaHistory_Weekly", 108#), _
        Array("ekaHistory_Monthly", 98#), _
        Array("ekaHistory_Overall", 90#), _
        Array("ekaHistory_Back", 86#))
    topPos = ws.Rows(3).Top + 4
    leftPos = ws.Range("A3").Left + 4
    For i = LBound(specs) To UBound(specs)
        Set shape = Nothing
        On Error Resume Next
        Set shape = ws.Shapes(CStr(specs(i)(0)))
        On Error GoTo 0
        If Not shape Is Nothing Then
            buttonWidth = CDbl(specs(i)(1))
            shape.Left = leftPos
            shape.Top = topPos
            shape.Width = buttonWidth
            shape.Height = 24
            shape.ZOrder 0
            leftPos = leftPos + buttonWidth + 7
        End If
    Next i
End Sub

Private Function ExistingReportSheet() As Worksheet
    On Error Resume Next
    Set ExistingReportSheet = ThisWorkbook.Worksheets(REPORT_SHEET)
    On Error GoTo 0
End Function

Private Function ResolveReportTool(ByVal fileName As String) As String
    Dim candidate As String, rootPath As String
    rootPath = ThisWorkbook.Path
    candidate = rootPath & "\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveReportTool = candidate: Exit Function
    candidate = rootPath & "\\..\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveReportTool = candidate: Exit Function
    candidate = rootPath & "\\..\\..\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveReportTool = candidate: Exit Function
End Function

Private Function QuoteReportArg(ByVal value As String) As String
    QuoteReportArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function ReadReportUtf8File(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadReportUtf8File = stream.ReadText
    stream.Close
End Function
'''


def _delete_shape(ws, name: str) -> None:
    try:
        ws.Shapes(name).Delete()
    except Exception:
        pass


def _add_button(ws, *, name: str, caption: str, macro: str, left: float, top: float, width: float) -> None:
    _delete_shape(ws, name)
    shape = ws.Shapes.AddShape(5, left, top, width, 24)
    shape.Name = name
    shape.OnAction = macro
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.Line.Weight = 1
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    try:
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    except Exception:
        pass
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3


def _last_used_row(ws) -> int:
    try:
        used = ws.UsedRange
        return max(1, int(used.Row) + int(used.Rows.Count) - 1)
    except Exception:
        return 1


def install_therapist_history_report(vbproject, workbook) -> None:
    try:
        component = vbproject.VBComponents(THERAPIST_HISTORY_MODULE_NAME)
        vbproject.VBComponents.Remove(component)
    except Exception:
        pass
    module = vbproject.VBComponents.Add(1)
    module.Name = THERAPIST_HISTORY_MODULE_NAME
    module.CodeModule.AddFromString(THERAPIST_HISTORY_MODULE_CODE)

    try:
        report_ws = workbook.Worksheets(THERAPIST_HISTORY_SHEET_NAME)
    except Exception:
        report_ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        report_ws.Name = THERAPIST_HISTORY_SHEET_NAME
    try:
        report_ws.Unprotect()
    except Exception:
        pass
    report_ws.Visible = -1
    report_ws.Cells.ClearContents()

    button_specs = (
        ("Daily", "Ημερήσια", "ShowTherapistHistoryDaily"),
        ("Weekly", "Εβδομαδιαία", "ShowTherapistHistoryWeekly"),
        ("Monthly", "Μηνιαία", "ShowTherapistHistoryMonthly"),
        ("Overall", "Συνολικά", "ShowTherapistHistoryOverall"),
        ("Back", "← MASTER", "GoToMaster"),
    )
    left = 4.0
    for suffix, caption, macro in button_specs:
        width = 90.0 if suffix != "Back" else 82.0
        _add_button(report_ws, name=THERAPIST_HISTORY_BUTTON_PREFIX + suffix, caption=caption, macro=macro, left=left, top=48.0, width=width)
        left += width + 6.0

    try:
        source_ws = workbook.Worksheets("THERAPIST_DAILY")
    except Exception as exc:
        raise RuntimeError("Workbook has no THERAPIST_DAILY sheet") from exc
    last_row = _last_used_row(source_ws)
    anchor = source_ws.Range(f"A{last_row + 2}")
    _add_button(source_ws, name=THERAPIST_HISTORY_NAV_BUTTON, caption="Στατιστικά θεραπευτών", macro=THERAPIST_HISTORY_OPEN_MACRO, left=float(anchor.Left) + 164, top=float(anchor.Top) + 2, width=126)
