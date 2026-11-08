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

Public Sub ShowSelectedTherapistChart()
    Dim ws As Worksheet, selectedRow As Long, finalRow As Long
    Dim priorEvents As Boolean, priorScreenUpdating As Boolean
    Dim failureNumber As Long, failureText As String
    priorEvents = Application.EnableEvents
    priorScreenUpdating = Application.ScreenUpdating
    On Error GoTo ChartError
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    If Not ActiveSheet Is ws Then
        MsgBox "Ανοίξτε πρώτα τα στατιστικά θεραπευτών.", vbExclamation
        Exit Sub
    End If
    selectedRow = ActiveCell.Row
    finalRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If selectedRow < 23 Or selectedRow > finalRow Then
        MsgBox "Επιλέξτε πρώτα τη γραμμή του θεραπευτή στον πίνακα.", vbInformation
        Exit Sub
    End If
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Unprotect ""
    ws.Range("P1").Value2 = selectedRow
    BuildTherapistLoadChart ws, CLng(ws.Cells(selectedRow, 5).Value2), _
        CLng(ws.Cells(selectedRow, 6).Value2), CStr(ws.Cells(selectedRow, 3).Value2), _
        CLng(ws.Cells(selectedRow, 7).Value2), CLng(ws.Cells(selectedRow, 8).Value2), _
        CLng(ws.Cells(selectedRow, 4).Value2), CDbl(ws.Cells(selectedRow, 14).Value2), selectedRow
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorScreenUpdating
    Exit Sub
ChartError:
    failureNumber = Err.Number
    failureText = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorScreenUpdating
    On Error GoTo 0
    MsgBox "Το ατομικό γράφημα δεν ενημερώθηκε (" & CStr(failureNumber) & _
        "): " & failureText, vbCritical, "Στατιστικά θεραπευτών"
End Sub

Public Sub HighlightSelectedTherapistRow(ByVal selectedRow As Long)
    Dim ws As Worksheet, finalRow As Long
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    finalRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If selectedRow < 23 Or selectedRow > finalRow Then Exit Sub
    ws.Unprotect ""
    ws.Range("P1").Value2 = selectedRow
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub

Private Sub RenderTherapistHistoryReport(ByVal reportMode As String)
    Dim exportScript As String, outputPath As String, trendPath As String, commandLine As String
    Dim exitCode As Long, rawText As String, lines As Variant, parts As Variant
    Dim trendLines As Variant, trendParts As Variant, trendLine As Long, trendRow As Long
    Dim ws As Worksheet, rowIndex As Long, sourceLine As Long, col As Long
    Dim headers As Variant
    Dim totalPlannedSlots As Long, totalActualSlots As Long
    Dim totalPlannedSessions As Long, totalActualSessions As Long
    Dim totalFinalizedTherapistDays As Long
    Dim totalDistinctPatientDayLoad As Double

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
    trendPath = Environ$("TEMP") & "\\eka_therapist_daily_trend_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".tsv"
    commandLine = QuoteReportArg("python") & " " & QuoteReportArg(exportScript) & _
                  " --workbook " & QuoteReportArg(ThisWorkbook.FullName) & _
                  " --output " & QuoteReportArg(outputPath) & _
                  " --mode " & QuoteReportArg(reportMode) & _
                  " --trend-output " & QuoteReportArg(trendPath)

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If exitCode <> 0 Or Dir$(outputPath) = "" Or Dir$(trendPath) = "" Then
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
    ws.Range("A1:P1048576").ClearContents
    ws.Range("P1").Value2 = 0
    ws.Range("P2").Value2 = reportMode
    EnsureReportButtons ws
    ws.Cells(1, 1).Value2 = "Ιστορικό / Στατιστικά Θεραπευτών"
    ws.Cells(2, 1).Value2 = ReportModeDescription(reportMode) & _
        " | Οριστικοποιημένες ημέρες | Σύγκριση με μέσο όρο ίδιας περιόδου"
    For col = 0 To UBound(headers)
        ws.Cells(22, col + 1).Value2 = headers(col)
    Next col

    rowIndex = 23
    For sourceLine = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(sourceLine)))) = 0 Then GoTo NextLine
        parts = Split(CStr(lines(sourceLine)), vbTab)
        If UBound(parts) <> 14 Then Err.Raise vbObjectError + 7601, , "Invalid therapist-history UI row"
        ws.Cells(rowIndex, 1).Value2 = DisplayPeriodStart(reportMode, parts(0), parts(1))
        ws.Cells(rowIndex, 2).Value2 = DisplayPeriodEnd(reportMode, parts(0), parts(1))
        ws.Cells(rowIndex, 3).Value2 = parts(3)
        ws.Cells(rowIndex, 4).Value2 = parts(4)
        For col = 5 To 12
            ws.Cells(rowIndex, col).Value2 = parts(col)
        Next col
        If Len(parts(12)) > 0 Then ws.Cells(rowIndex, 12).Value2 = parts(12) & "%"
        If Len(parts(13)) > 0 Then ws.Cells(rowIndex, 13).Value2 = parts(13) & "%"
        If Len(parts(14)) > 0 Then
            ws.Cells(rowIndex, 14).Value2 = CDbl(Replace(parts(14), ".", Application.DecimalSeparator))
            ws.Cells(rowIndex, 14).Interior.Color = _
                WorkloadTrafficColor(CDbl(ws.Cells(rowIndex, 14).Value2))
            ws.Cells(rowIndex, 14).Font.Bold = True
            ws.Cells(rowIndex, 14).NumberFormat = "0.00"
        End If
        totalDistinctPatientDayLoad = totalDistinctPatientDayLoad + _
            CDbl(ws.Cells(rowIndex, 14).Value2) * CLng(parts(4))
        totalPlannedSlots = totalPlannedSlots + CLng(parts(5))
        totalActualSlots = totalActualSlots + CLng(parts(6))
        totalPlannedSessions = totalPlannedSessions + CLng(parts(7))
        totalActualSessions = totalActualSessions + CLng(parts(8))
        totalFinalizedTherapistDays = totalFinalizedTherapistDays + CLng(parts(4))
        rowIndex = rowIndex + 1
NextLine:
    Next sourceLine

    ' Cache only daily finalized facts for chart drawing. Q:U are hidden,
    ' sorted by date from the exporter, and cleared on every fresh render.
    trendLines = Split(Replace(Replace( _
        ReadReportUtf8File(trendPath), ChrW(&HFEFF), ""), vbCr, ""), vbLf)
    If UBound(trendLines) < 0 Or CStr(trendLines(0)) <> _
            "Date" & vbTab & "Therapist" & vbTab & "Planned" & vbTab & _
            "Delivered" & vbTab & "UniquePatients" Then
        Err.Raise vbObjectError + 7602, , "Invalid therapist daily trend header"
    End If
    trendRow = 2
    For trendLine = 1 To UBound(trendLines)
        If Len(Trim$(CStr(trendLines(trendLine)))) = 0 Then GoTo NextTrendLine
        trendParts = Split(CStr(trendLines(trendLine)), vbTab)
        If UBound(trendParts) <> 4 Then Err.Raise vbObjectError + 7603, , _
            "Invalid therapist daily trend record"
        ws.Cells(trendRow, 17).Value2 = trendParts(0)
        ws.Cells(trendRow, 18).Value2 = trendParts(1)
        ws.Cells(trendRow, 19).Value2 = CLng(trendParts(2))
        ws.Cells(trendRow, 20).Value2 = CLng(trendParts(3))
        ws.Cells(trendRow, 21).Value2 = CLng(trendParts(4))
        trendRow = trendRow + 1
NextTrendLine:
    Next trendLine

    FormatReportSheet ws, rowIndex - 1
    BuildTherapistLoadChart ws, totalPlannedSlots, totalActualSlots, "", _
        totalPlannedSessions, totalActualSessions, totalFinalizedTherapistDays, _
        totalDistinctPatientDayLoad / Application.Max(1, totalFinalizedTherapistDays), 0
    EnsureReportButtons ws
    HighlightTherapistReportMode ws, reportMode
    ws.Activate
    KeepApplicationShell
    ResetReportWindow ws
    EnsureReportButtons ws
    HighlightTherapistReportMode ws, reportMode
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True

CleanUp:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    If Len(outputPath) > 0 And Dir$(outputPath) <> "" Then Kill outputPath
    If Len(trendPath) > 0 And Dir$(trendPath) <> "" Then Kill trendPath
    On Error GoTo 0
    Exit Sub

ReportError:
    MsgBox "Δεν ήταν δυνατή η ενημέρωση του ιστορικού θεραπευτών.", vbCritical, "Ιστορικό θεραπευτών"
    Resume CleanUp
End Sub

Private Function ReportHeaders(ByVal reportMode As String) As Variant
    Select Case LCase$(Trim$(reportMode))
        Case "daily"
            ReportHeaders = Array("Ημερομηνία", "", "Θεραπευτής", "Οριστ. ημέρες", "Προγρ. θέσεις", "Πραγμ. θέσεις", "Προγρ. συνεδρίες", "Πραγμ. συνεδρίες", "Αναπληρώσεις", "Απουσίες/αναβολές", "Αναστολές θεραπείας", "Προγρ. φόρτος", "Πραγμ. φόρτος", "Ασθενείς/ημέρα")
        Case "monthly"
            ReportHeaders = Array("Μήνας", "", "Θεραπευτής", "Οριστ. ημέρες", "Προγρ. θέσεις", "Πραγμ. θέσεις", "Προγρ. συνεδρίες", "Πραγμ. συνεδρίες", "Αναπληρώσεις", "Απουσίες/αναβολές", "Αναστολές θεραπείας", "Προγρ. φόρτος", "Πραγμ. φόρτος", "Ασθενείς/ημέρα")
        Case Else
            ReportHeaders = Array("Από", "Έως", "Θεραπευτής", "Οριστ. ημέρες", "Προγρ. θέσεις", "Πραγμ. θέσεις", "Προγρ. συνεδρίες", "Πραγμ. συνεδρίες", "Αναπληρώσεις", "Απουσίες/αναβολές", "Αναστολές θεραπείας", "Προγρ. φόρτος", "Πραγμ. φόρτος", "Ασθενείς/ημέρα")
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

Private Function WorkloadTrafficColor(ByVal loadPerDay As Double) As Long
    ' Classification applies to distinct patients treated per finalized
    ' therapist-day. Session count and occupied slots are separate KPIs.
    If loadPerDay >= 5.5 Then
        WorkloadTrafficColor = RGB(255, 3, 3)        ' TRIAL #FF0303: six or more
    ElseIf loadPerDay >= 4.5 Then
        WorkloadTrafficColor = RGB(255, 180, 0)      ' TRIAL #FFB400: five
    ElseIf loadPerDay >= 2.5 Then
        WorkloadTrafficColor = RGB(255, 242, 0)      ' TRIAL #FFF200: three/four
    ElseIf loadPerDay > 0 Then
        WorkloadTrafficColor = RGB(91, 250, 116)     ' TRIAL #5BFA74: one/two
    Else
        WorkloadTrafficColor = RGB(63, 168, 245)     ' TRIAL #3FA8F5: zero
    End If
End Function

Private Function NewTherapistMiniChart(ByVal ws As Worksheet, _
        ByVal chartName As String, ByVal firstColumn As String, _
        ByVal lastColumn As String) As ChartObject
    Dim plot As ChartObject, chartWidth As Double
    chartWidth = ws.Range(lastColumn & "5").Left + _
        ws.Range(lastColumn & "5").Width - ws.Range(firstColumn & "5").Left
    Set plot = ws.ChartObjects.Add( _
        ws.Range(firstColumn & "5").Left, ws.Range(firstColumn & "5").Top, _
        chartWidth - 4, 108)
    plot.Name = chartName
    plot.Placement = xlFreeFloating
    plot.Chart.ChartType = xlColumnClustered
    plot.Chart.HasLegend = False
    Set NewTherapistMiniChart = plot
End Function

Private Sub AddTherapistMeanMarker(ByVal plot As ChartObject, _
        ByVal markerName As String, ByVal meanValue As Double, _
        ByVal categoryIndex As Long, ByVal categoryCount As Long, _
        ByVal markerColor As Long)
    ' Chart-local shape gives a genuinely SHORT, thin, solid horizontal
    ' reference across the column, not a full-width line-series or a bar.
    ' Category slot determines x; value-axis determines y.
    Dim area As PlotArea, axis As Axis, marker As Shape
    Dim slotWidth As Double, centerX As Double, lineHalfWidth As Double
    Dim markerY As Double, minValue As Double, maxValue As Double

    If meanValue < 0 Or categoryCount <= 0 Then Exit Sub
    Set axis = plot.Chart.Axes(xlValue)
    minValue = CDbl(axis.MinimumScale)
    maxValue = CDbl(axis.MaximumScale)
    If maxValue <= minValue Then Exit Sub
    Set area = plot.Chart.PlotArea
    slotWidth = CDbl(area.InsideWidth) / CDbl(categoryCount)
    centerX = CDbl(area.InsideLeft) + _
        (CDbl(categoryIndex) - 0.5) * slotWidth
    ' Marker spans about half a category slot, slightly wider than the
    ' plotted column on a one-series clustered column chart.
    lineHalfWidth = slotWidth * 0.28
    markerY = CDbl(area.InsideTop) + CDbl(area.InsideHeight) * _
        (maxValue - meanValue) / (maxValue - minValue)
    Set marker = plot.Chart.Shapes.AddLine( _
        CSng(centerX - lineHalfWidth), CSng(markerY), _
        CSng(centerX + lineHalfWidth), CSng(markerY))
    marker.Name = markerName
    marker.Line.ForeColor.RGB = markerColor
    marker.Line.Weight = 1.25
    marker.Line.DashStyle = msoLineSolid
End Sub

Private Sub BuildTherapistLoadChart(ByVal ws As Worksheet, _
        ByVal plannedSlots As Long, ByVal actualSlots As Long, _
        Optional ByVal therapistName As String = "", _
        Optional ByVal plannedSessions As Long = 0, _
        Optional ByVal actualSessions As Long = 0, _
        Optional ByVal finalizedDays As Long = 0, _
        Optional ByVal uniquePatientsPerDay As Double = 0, _
        Optional ByVal selectedRow As Long = 0)
    Dim plot As ChartObject, series As Series
    Dim loadPerDay As Double, referencePatients As Double
    Dim referencePlanned As Double, referenceActual As Double
    Dim teamPlanned As Long, teamActual As Long
    Dim referenceColor As Long
    Dim cohortMaxPatients As Double, cohortMaxSessions As Double
    Dim cohortAxisPatients As Double, cohortAxisSessions As Double
    Dim referenceCount As Long, referenceDays As Long, r As Long
    Dim lastRow As Long, selectedPeriod As String, periodKey As String
    Dim reportMode As String

    ' Always remove all old charts, including when there are no rows.
    On Error Resume Next
    ws.ChartObjects("ekaTherapistLoadChart").Delete
    ws.ChartObjects("ekaTherapistSessionChart").Delete
    ws.ChartObjects("ekaTherapistAttendanceChart").Delete
    ws.ChartObjects("ekaTherapistLoadTrend").Delete
    ws.ChartObjects("ekaTherapistSessionTrend").Delete
    On Error GoTo 0
    ws.Range("A5:M20").ClearContents

    If finalizedDays <= 0 Then
        ws.Range("A6").Value2 = "Δεν υπάρχουν οριστικοποιημένες ημέρες για γράφημα."
        Exit Sub
    End If

    ' Comparable, period-specific benchmark. Do not compare September's
    ' therapist against October's or mix weekly and monthly rows.
    ' One row is one therapist-period; load is weighted by finalized days.
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    reportMode = CStr(ws.Range("P2").Value2)
    If selectedRow >= 23 And selectedRow <= lastRow Then
        selectedPeriod = CStr(ws.Cells(selectedRow, 1).Value2) & "|" & _
                         CStr(ws.Cells(selectedRow, 2).Value2)
    Else
        selectedPeriod = ""
    End If
    For r = 23 To lastRow
        periodKey = CStr(ws.Cells(r, 1).Value2) & "|" & CStr(ws.Cells(r, 2).Value2)
        If selectedRow = 0 Or periodKey = selectedPeriod Then
            If IsNumeric(ws.Cells(r, 14).Value2) And _
               IsNumeric(ws.Cells(r, 4).Value2) Then
                referenceDays = referenceDays + CLng(ws.Cells(r, 4).Value2)
                referencePatients = referencePatients + _
                    CDbl(ws.Cells(r, 14).Value2) * CLng(ws.Cells(r, 4).Value2)
                referencePlanned = referencePlanned + CDbl(ws.Cells(r, 7).Value2)
                referenceActual = referenceActual + CDbl(ws.Cells(r, 8).Value2)
                ' Chart scales are based on the WHOLE comparable cohort.
                ' Selecting a person must never rescale the ordinate.
                cohortMaxPatients = Application.Max(cohortMaxPatients, _
                    CDbl(ws.Cells(r, 14).Value2))
                cohortMaxSessions = Application.Max(cohortMaxSessions, _
                    CDbl(ws.Cells(r, 7).Value2), CDbl(ws.Cells(r, 8).Value2))
                referenceCount = referenceCount + 1
            End If
        End If
    Next r
    ' One visually rounded shared ceiling per selected period/cohort.
    ' Include team benchmark in case all valid rows are edge cases.
    cohortAxisPatients = Application.Max(6, _
        WorksheetFunction.RoundUp(cohortMaxPatients, 0))
    cohortAxisSessions = Application.Max(1, _
        WorksheetFunction.RoundUp(cohortMaxSessions, -1))
    teamPlanned = CLng(referencePlanned)
    teamActual = CLng(referenceActual)
    If referenceCount > 0 And referenceDays > 0 Then
        referencePatients = referencePatients / referenceDays
        referencePlanned = referencePlanned / referenceCount
        referenceActual = referenceActual / referenceCount
    End If

    ' One plain-language departmental total, same period as the benchmark.
    ' A single combined total is not a sum of therapist percentage values.
    If referenceCount > 0 Then
        ws.Range("A20").Value2 = "ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ | Προγραμματισμένες: " & _
            CStr(teamPlanned) & " | Πραγματοποιημένες: " & CStr(teamActual) & _
            " | Υλοποίηση: " & IIf(teamPlanned > 0, _
            Format$(100# * CDbl(teamActual) / Application.Max(1, teamPlanned), "0.0") & "%", "—")
        ws.Range("A20").Font.Bold = True
        ws.Range("A20").Font.Color = RGB(37, 99, 170)
    End If

    ' Chart one: 0-6 delivered treatments per finalized therapist-day,
    ' with explicit traffic colors; greater than six is still red.
    loadPerDay = uniquePatientsPerDay
    Set plot = NewTherapistMiniChart(ws, "ekaTherapistLoadChart", "A", "D")
    Set series = plot.Chart.SeriesCollection.NewSeries
    series.Name = "Θεραπευτής"
    series.XValues = Array("Ανά ημέρα")
    series.Values = Array(loadPerDay)
    series.Format.Fill.ForeColor.RGB = WorkloadTrafficColor(loadPerDay)
    plot.Chart.HasLegend = False
    plot.Chart.Axes(xlValue).MinimumScale = 0
    plot.Chart.Axes(xlValue).MaximumScale = cohortAxisPatients
    ws.Range("A19").Value2 = "Μ.Ο. περιόδου: " & Format$(referencePatients, "0.0") & " ασθενείς/ημέρα"
    plot.Chart.HasTitle = True
    plot.Chart.ChartTitle.Text = "Ασθενείς ανά ημέρα" & IIf( _
        Len(therapistName) > 0, " | " & therapistName, "")
    referenceColor = RGB(225, 0, 30) ' red by default
    If WorkloadTrafficColor(loadPerDay) = RGB(255, 3, 3) Then
        referenceColor = RGB(57, 255, 20) ' fluorescent green on red workload bar
    End If
    AddTherapistMeanMarker plot, "ekaMeanPatients", referencePatients, 1, 1, referenceColor

    ' Chart two: planned/delivered treatment sessions per reporting period.
    Set plot = NewTherapistMiniChart(ws, "ekaTherapistSessionChart", "E", "H")
    Set series = plot.Chart.SeriesCollection.NewSeries
    series.Name = "Θεραπευτής"
    series.XValues = Array("Προγραμ.", "Πραγματ.")
    series.Values = Array(plannedSessions, actualSessions)
    ' User-selected semantic color contract: BLUE = planned, RED = delivered.
    ' Apply the same mapping to the mini time-series beneath this chart.
    series.Points(1).Format.Fill.ForeColor.RGB = RGB(63, 168, 245)
    series.Points(2).Format.Fill.ForeColor.RGB = RGB(255, 3, 3)
    plot.Chart.HasLegend = False
    plot.Chart.Axes(xlValue).MinimumScale = 0
    plot.Chart.Axes(xlValue).MaximumScale = cohortAxisSessions
    ws.Range("E19").Value2 = "Κόκκινοι δείκτες: Μ.Ο. προγρ. " & _
        Format$(referencePlanned, "0.0") & " | πραγματ. " & _
        Format$(referenceActual, "0.0")
    plot.Chart.HasTitle = True
    plot.Chart.ChartTitle.Text = "Θεραπείες" & IIf( _
        Len(therapistName) > 0, " | " & therapistName, "")
    AddTherapistMeanMarker plot, "ekaMeanPlanned", referencePlanned, 1, 2, RGB(225, 0, 30)
    AddTherapistMeanMarker plot, "ekaMeanActual", referenceActual, 2, 2, RGB(57, 255, 20)

    ' Mini trend panels: genuine finalized daily data, same time window and
    ' chart ordinate for every therapist in that cohort, two series in sessions.
    BuildTherapistMiniTrends ws, reportMode, selectedRow, therapistName

    ' Chart three: not a zero-absence chart. Historical therapist attendance
    ' and a reliable working-day denominator are NOT persisted in the current
    ' finalized therapist/day history schema. Do not manufacture that count.
    Set plot = NewTherapistMiniChart(ws, "ekaTherapistAttendanceChart", "I", "M")
    plot.Chart.HasTitle = True
    plot.Chart.ChartTitle.Text = "Παρουσίες / απουσίες"
    ws.Range("I19").Value2 = "Μη διαθέσιμο ιστορικό απουσιών / εργασίμων"
    ws.Range("I19").Font.Bold = True
End Sub

Private Sub FormatReportSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    On Error Resume Next
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Range("A1:M1").Merge
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:M2").Merge
    ws.Range("A2").Font.Bold = True
    ws.Range("A22:N22").Font.Bold = True
    ws.Range("A22:N22").WrapText = True
    ws.Columns("A:B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 19
    ws.Columns("D:N").ColumnWidth = 13
    If lastRow >= 22 Then ws.Range("A22:N" & CStr(lastRow)).Borders.LineStyle = 1
    ws.Range("A23:M" & CStr(Application.Max(lastRow, 23))).FormatConditions.Delete
    If lastRow >= 23 Then
        With ws.Range("A23:M" & CStr(lastRow)).FormatConditions.Add( _
                Type:=xlExpression, Formula1:="=ROW()=$P$1")
            .Interior.Color = RGB(37, 99, 170)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
    End If
    ws.Columns("P:U").Hidden = True
    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 6
    ws.Rows(22).RowHeight = 30
    PositionReportButtons ws
    ws.ScrollArea = "A1:N" & CStr(Application.Max(lastRow + 2, 25))
    On Error GoTo 0
End Sub

Private Sub ResetReportWindow(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = 0
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        ws.Range("A23").Select
        ActiveWindow.SplitRow = 22
        ActiveWindow.FreezePanes = True
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    End If
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Sub HighlightTherapistReportMode(ByVal ws As Worksheet, ByVal reportMode As String)
    Dim modes As Variant, i As Long, shape As Shape, selected As String
    modes = Array("Daily", "Weekly", "Monthly", "Overall")
    Select Case LCase$(Trim$(reportMode))
        Case "daily": selected = "Daily"
        Case "weekly": selected = "Weekly"
        Case "monthly": selected = "Monthly"
        Case "overall": selected = "Overall"
        Case Else: Err.Raise vbObjectError + 7650, , "Unknown therapist mode"
    End Select
    For i = LBound(modes) To UBound(modes)
        Set shape = ws.Shapes("ekaHistory_" & CStr(modes(i)))
        If CStr(modes(i)) = selected Then
            shape.Fill.ForeColor.RGB = RGB(37, 99, 170)
            shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            shape.Line.ForeColor.RGB = RGB(18, 60, 115)
            shape.Line.Weight = 2
        Else
            shape.Fill.ForeColor.RGB = RGB(236, 241, 247)
            shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            shape.Line.ForeColor.RGB = RGB(170, 185, 205)
            shape.Line.Weight = 1
        End If
    Next i
End Sub

Private Sub EnsureReportButtons(ByVal ws As Worksheet)
    EnsureOneReportButton ws, "ekaHistory_Daily", "Ημερήσια", "ShowTherapistHistoryDaily", 98#
    EnsureOneReportButton ws, "ekaHistory_Weekly", "Εβδομαδιαία", "ShowTherapistHistoryWeekly", 108#
    EnsureOneReportButton ws, "ekaHistory_Monthly", "Μηνιαία", "ShowTherapistHistoryMonthly", 98#
    EnsureOneReportButton ws, "ekaHistory_Overall", "Συνολικά", "ShowTherapistHistoryOverall", 90#
    EnsureOneReportButton ws, "ekaHistory_Back", "← Προηγούμενη", "ToolbarTherapistDaily", 116#
    EnsureOneReportButton ws, "ekaHistory_PersonChart", "Γράφημα θεραπευτή", "ShowSelectedTherapistChart", 160#
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
    shape.Visible = True
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
        Array("ekaHistory_Back", 116#), _
        Array("ekaHistory_PersonChart", 160#))
    topPos = ws.Rows(3).Top + 4
    leftPos = ws.Range("A3").Left + 4
    For i = LBound(specs) To UBound(specs)
        Set shape = Nothing
        On Error Resume Next
        Set shape = ws.Shapes(CStr(specs(i)(0)))
        On Error GoTo 0
        If Not shape Is Nothing Then
            buttonWidth = CDbl(specs(i)(1))
            shape.Visible = True
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
    # Clear any prior worksheet protection explicitly before rebuilding the report UI.
    try:
        report_ws.Unprotect("")
    except Exception as exc:
        if bool(report_ws.ProtectContents) or bool(report_ws.ProtectDrawingObjects):
            raise RuntimeError(
                "Cannot unprotect therapist-history report before installation "
                f"(ProtectContents={bool(report_ws.ProtectContents)}, "
                f"ProtectDrawingObjects={bool(report_ws.ProtectDrawingObjects)})"
            ) from exc
    if bool(report_ws.ProtectContents) or bool(report_ws.ProtectDrawingObjects):
        raise RuntimeError(
            "Therapist-history report remained protected after Unprotect "
            f"(ProtectContents={bool(report_ws.ProtectContents)}, "
            f"ProtectDrawingObjects={bool(report_ws.ProtectDrawingObjects)})"
        )
    report_ws.Visible = -1
    report_ws.Cells.ClearContents()

    button_specs = (
        ("Daily", "Ημερήσια", "ShowTherapistHistoryDaily"),
        ("Weekly", "Εβδομαδιαία", "ShowTherapistHistoryWeekly"),
        ("Monthly", "Μηνιαία", "ShowTherapistHistoryMonthly"),
        ("Overall", "Συνολικά", "ShowTherapistHistoryOverall"),
        ("Back", "← Προηγούμενη", "ToolbarTherapistDaily"),
    )
    left = 4.0
    for suffix, caption, macro in button_specs:
        width = 90.0 if suffix != "Back" else 116.0
        _add_button(report_ws, name=THERAPIST_HISTORY_BUTTON_PREFIX + suffix, caption=caption, macro=macro, left=left, top=48.0, width=width)
        left += width + 6.0

    try:
        source_ws = workbook.Worksheets("THERAPIST_DAILY")
    except Exception as exc:
        raise RuntimeError("Workbook has no THERAPIST_DAILY sheet") from exc
    last_row = _last_used_row(source_ws)
    anchor = source_ws.Range(f"A{last_row + 2}")
    _add_button(source_ws, name=THERAPIST_HISTORY_NAV_BUTTON, caption="Στατιστικά θεραπευτών", macro=THERAPIST_HISTORY_OPEN_MACRO, left=float(anchor.Left) + 164, top=float(anchor.Top) + 2, width=126)

    # Keep the blank report surface protected from direct edits immediately after install.
    # DrawingObjects=False preserves the report navigation buttons.
    # Use positional COM arguments to avoid any named-argument dispatch ambiguity.
    report_ws.Protect("", False, True, True)
    protect_contents = bool(report_ws.ProtectContents)
    protect_objects = bool(report_ws.ProtectDrawingObjects)
    if not protect_contents or protect_objects:
        raise RuntimeError(
            "Therapist-history report protection state is invalid after Protect "
            f"(ProtectContents={protect_contents}, ProtectDrawingObjects={protect_objects})"
        )
