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

    headers = Array("Από", "Έως", "Έκδοση", "Θεραπευτής", "Ημέρες", "Προγρ. θέσεις", "Πραγμ. θέσεις", "Προγρ. συνεδρίες", "Πραγμ. συνεδρίες", "Αναπληρώσεις", "Αποδεσμεύσεις", "Αναστολές", "Προγρ. φόρτος", "Πραγμ. φόρτος")

    ws.Unprotect
    ws.Range("A1:N1048576").ClearContents
    ws.Cells(1, 1).Value2 = "Ιστορικό / Στατιστικά Θεραπευτών"
    ws.Cells(2, 1).Value2 = "Προβολή: " & ReportModeCaption(reportMode)
    For col = 0 To UBound(headers)
        ws.Cells(4, col + 1).Value2 = headers(col)
    Next col

    rowIndex = 5
    For sourceLine = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(sourceLine)))) = 0 Then GoTo NextLine
        parts = Split(CStr(lines(sourceLine)), vbTab)
        If UBound(parts) <> 13 Then Err.Raise vbObjectError + 7601, , "Invalid therapist-history UI row"
        For col = 0 To 11
            ws.Cells(rowIndex, col + 1).Value2 = parts(col)
        Next col
        If Len(parts(12)) > 0 Then ws.Cells(rowIndex, 13).Value2 = parts(12) & "%"
        If Len(parts(13)) > 0 Then ws.Cells(rowIndex, 14).Value2 = parts(13) & "%"
        rowIndex = rowIndex + 1
NextLine:
    Next sourceLine

    FormatReportSheet ws, rowIndex - 1
    ws.Protect
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

Private Sub FormatReportSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    On Error Resume Next
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Range("A1:N1").Merge
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:N2").Merge
    ws.Range("A2").Font.Bold = True
    ws.Range("A4:N4").Font.Bold = True
    ws.Range("A4:N4").WrapText = True
    ws.Columns("A:B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E:N").ColumnWidth = 13
    If lastRow >= 4 Then ws.Range("A4:N" & CStr(lastRow)).Borders.LineStyle = 1
    ws.Rows("1:4").RowHeight = 22
    ws.ScrollArea = "A1:N" & CStr(Application.Max(lastRow + 2, 12))
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.FreezePanes = False
        ws.Range("A5").Select
        ActiveWindow.SplitRow = 4
        ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
End Sub

Private Function ExistingReportSheet() As Worksheet
    On Error Resume Next
    Set ExistingReportSheet = ThisWorkbook.Worksheets(REPORT_SHEET)
    On Error GoTo 0
End Function

Private Function ReportModeCaption(ByVal reportMode As String) As String
    Select Case LCase$(Trim$(reportMode))
        Case "daily": ReportModeCaption = "Ημερήσια"
        Case "weekly": ReportModeCaption = "Εβδομαδιαία"
        Case "monthly": ReportModeCaption = "Μηνιαία"
        Case "overall": ReportModeCaption = "Συνολικά"
        Case Else: ReportModeCaption = reportMode
    End Select
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
        _add_button(
            report_ws,
            name=THERAPIST_HISTORY_BUTTON_PREFIX + suffix,
            caption=caption,
            macro=macro,
            left=left,
            top=48.0,
            width=width,
        )
        left += width + 6.0

    try:
        source_ws = workbook.Worksheets("THERAPIST_DAILY")
    except Exception as exc:
        raise RuntimeError("Workbook has no THERAPIST_DAILY sheet") from exc
    last_row = _last_used_row(source_ws)
    anchor = source_ws.Range(f"A{last_row + 2}")
    _add_button(
        source_ws,
        name=THERAPIST_HISTORY_NAV_BUTTON,
        caption="Στατιστικά θεραπευτών",
        macro=THERAPIST_HISTORY_OPEN_MACRO,
        left=float(anchor.Left) + 164,
        top=float(anchor.Top) + 2,
        width=126,
    )
