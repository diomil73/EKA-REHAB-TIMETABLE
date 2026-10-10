from __future__ import annotations


MODULE_NAME = "modOutpatientInSession"

MODULE_CODE = r'''Option Explicit

Public Function SaveOutpatientScheduleInWorkbook( _
    ByVal patientId As String, _
    ByVal treatment As String, _
    ByVal timeText As String, _
    ByVal dayPattern As String, _
    ByVal therapistName As String _
) As String
    Dim ws As Worksheet
    Dim patients As Worksheet
    Dim patientRow As Long
    Dim patientName As String
    Dim patientTypeCol As Long
    Dim targetRow As Long
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo SaveError

    patientId = Trim$(patientId)
    treatment = Trim$(treatment)
    timeText = Trim$(timeText)
    dayPattern = Trim$(dayPattern)
    therapistName = Trim$(therapistName)

    If Len(patientId) = 0 Then Err.Raise vbObjectError + 2710, , "Το PatientID είναι υποχρεωτικό."
    If Len(treatment) = 0 Then Err.Raise vbObjectError + 2711, , "Η θεραπεία είναι υποχρεωτική."
    If Len(timeText) = 0 Or Not IsDate(Replace(timeText, ".", ":")) Then Err.Raise vbObjectError + 2712, , "Η ώρα δεν είναι έγκυρη."
    If Len(dayPattern) = 0 Then Err.Raise vbObjectError + 2713, , "Οι ημέρες είναι υποχρεωτικές."
    If Len(therapistName) = 0 Then Err.Raise vbObjectError + 2714, , "Ο θεραπευτής είναι υποχρεωτικός."

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set patients = ThisWorkbook.Worksheets("PATIENTS")
    patientRow = FindPatientRow(patients, patientId)
    If patientRow = 0 Then Err.Raise vbObjectError + 2715, , "Δεν βρέθηκε ο ασθενής " & patientId & "."

    patientName = Trim$(CStr(patients.Cells(patientRow, 3).Value))
    If Len(patientName) = 0 Then Err.Raise vbObjectError + 2716, , "Ο ασθενής δεν έχει όνομα."

    patientTypeCol = FindHeader(patients, Array("PatientType", "ΤύποςΑσθενή", "Τύπος Ασθενή"))
    If patientTypeCol > 0 Then
        If Not IsOutpatientValue(CStr(patients.Cells(patientRow, patientTypeCol).Value)) Then
            Err.Raise vbObjectError + 2717, , "Ο ασθενής δεν είναι καταχωρημένος ως εξωτερικός."
        End If
    End If

    Set ws = EnsureOutpatientScheduleSheet()
    targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2

    ws.Cells(targetRow, 1).Value = patientId
    ws.Cells(targetRow, 2).Value = patientName
    ws.Cells(targetRow, 3).Value = treatment
    ws.Cells(targetRow, 4).Value = TimeValue(Replace(timeText, ".", ":"))
    ws.Cells(targetRow, 4).NumberFormat = "hh:mm"
    ws.Cells(targetRow, 5).Value = dayPattern
    ws.Cells(targetRow, 6).Value = therapistName

    RefreshTherapistDailyForCurrentDate patientName, therapistName, timeText, dayPattern, treatment

    Application.CalculateFull
    ThisWorkbook.Save

    SaveOutpatientScheduleInWorkbook = "outpatient:" & CStr(targetRow)

SaveExit:
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Function

SaveError:
    Dim errorText As String
    errorText = Err.Description
    On Error Resume Next
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 2720, "SaveOutpatientScheduleInWorkbook", errorText
End Function

Private Sub RefreshTherapistDailyForCurrentDate( _
    ByVal patientName As String, _
    ByVal therapistName As String, _
    ByVal timeText As String, _
    ByVal dayPattern As String, _
    ByVal treatment As String _
)
    Dim dailyInput As Worksheet
    Dim dailySheet As Worksheet
    Dim targetDate As Date
    Dim targetCell As Range
    Dim previousText As String
    Dim lineStart As Long
    Dim linePosition As Long
    Dim normalizedTime As Date

    On Error GoTo RefreshError

    Set dailyInput = ThisWorkbook.Worksheets("DAILY_INPUT")
    If Not TryReadDailyInputDate(dailyInput.Range("B2").Value, targetDate) Then Exit Sub
    If Not DayPatternIncludesDate(dayPattern, targetDate) Then Exit Sub

    normalizedTime = TimeValue(Replace(timeText, ".", ":"))
    Set dailySheet = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    Set targetCell = FindTherapistDailyCell(dailySheet, therapistName, normalizedTime)
    If targetCell Is Nothing Then
        Err.Raise vbObjectError + 2730, , _
            "Δεν βρέθηκε θέση στο THERAPIST_DAILY για " & therapistName & " στις " & Format$(normalizedTime, "hh:mm") & "."
    End If

    previousText = CStr(targetCell.Value)
    linePosition = InStr(1, vbLf & previousText & vbLf, vbLf & patientName & vbLf, vbTextCompare)
    If linePosition = 0 Then
        If Len(Trim$(previousText)) = 0 Then
            targetCell.Value = patientName
            targetCell.Interior.Color = RGB(221, 235, 247)
            lineStart = 1
        Else
            targetCell.Value = previousText & vbLf & patientName
            lineStart = Len(previousText) + 2
        End If
    Else
        lineStart = linePosition
    End If

    targetCell.WrapText = True
    If lineStart > 0 Then
        If IsRoboticTreatment(treatment) Then
            targetCell.Characters(lineStart, Len(patientName)).Font.Color = RGB(255, 140, 0)
            targetCell.Characters(lineStart, Len(patientName)).Font.Bold = True
        Else
            targetCell.Characters(lineStart, Len(patientName)).Font.Color = RGB(0, 0, 0)
        End If
    End If
    Exit Sub

RefreshError:
    Err.Raise Err.Number, "RefreshTherapistDailyForCurrentDate", Err.Description
End Sub

Private Function TryReadDailyInputDate(ByVal rawValue As Variant, ByRef result As Date) As Boolean
    Dim text As String
    Dim parts() As String

    If IsDate(rawValue) Then
        result = CDate(rawValue)
        TryReadDailyInputDate = True
        Exit Function
    End If

    text = Trim$(CStr(rawValue))
    If Len(text) = 0 Then Exit Function
    parts = Split(text, "/")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then Exit Function

    On Error GoTo InvalidDate
    result = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
    TryReadDailyInputDate = True
    Exit Function

InvalidDate:
    TryReadDailyInputDate = False
End Function

Private Function DayPatternIncludesDate(ByVal dayPattern As String, ByVal targetDate As Date) As Boolean
    Dim text As String
    Dim token As String

    text = LCase$(Trim$(dayPattern))
    If Len(text) = 0 Then Exit Function

    If InStr(1, text, "καθ", vbTextCompare) > 0 Or InStr(1, text, "daily", vbTextCompare) > 0 Then
        DayPatternIncludesDate = True
        Exit Function
    End If

    Select Case Weekday(targetDate, vbMonday)
        Case 1: token = "δε"
        Case 2: token = "τρ"
        Case 3: token = "τε"
        Case 4: token = "πε"
        Case 5: token = "πα"
        Case 6: token = "σα"
        Case 7: token = "κυ"
    End Select

    DayPatternIncludesDate = (InStr(1, text, token, vbTextCompare) > 0)
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
    If IsDate(text) Then
        TimeKey = Format$(TimeValue(text), "hh:mm")
    End If
    Exit Function

InvalidTime:
    TimeKey = ""
End Function

Private Function FindTherapistDailyCell( _
    ByVal ws As Worksheet, _
    ByVal therapistName As String, _
    ByVal slotTime As Date _
) As Range
    Dim headerRows As Variant
    Dim firstRows As Variant
    Dim lastRows As Variant
    Dim blockIndex As Long
    Dim headerRow As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim providerCol As Long
    Dim col As Long
    Dim rowIndex As Long
    Dim valueText As String
    Dim targetProvider As String
    Dim targetTime As String

    headerRows = Array(1, 11)
    firstRows = Array(2, 12)
    lastRows = Array(8, 18)
    targetProvider = NormalizeProviderName(therapistName)
    targetTime = Format$(slotTime, "hh:mm")

    For blockIndex = 0 To 1
        headerRow = headerRows(blockIndex)
        firstRow = firstRows(blockIndex)
        lastRow = lastRows(blockIndex)
        providerCol = 0

        For col = 2 To 10
            valueText = CStr(ws.Cells(headerRow, col).Value)
            If NormalizeProviderName(valueText) = targetProvider Then
                providerCol = col
                Exit For
            End If
        Next col

        If providerCol > 0 Then
            For rowIndex = firstRow To lastRow
                If TimeKey(ws.Cells(rowIndex, 1).Value) = targetTime Or _
                   TimeKey(ws.Cells(rowIndex, 1).Text) = targetTime Then
                    Set FindTherapistDailyCell = ws.Cells(rowIndex, providerCol)
                    Exit Function
                End If
            Next rowIndex
        End If
    Next blockIndex
End Function

Private Function IsRoboticTreatment(ByVal treatment As String) As Boolean
    Dim text As String
    text = LCase$(Trim$(treatment))
    IsRoboticTreatment = (InStr(1, text, "ρομποτ", vbTextCompare) > 0 Or _
                          InStr(1, text, "robot", vbTextCompare) > 0)
End Function

Private Function EnsureOutpatientScheduleSheet() As Worksheet
    Dim ws As Worksheet
    Dim headers As Variant
    Dim index As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "OUTPATIENT_SCHEDULE"
    End If

    headers = Array("PatientID", "Ασθενής", "Θεραπεία", "Ώρα", "Ημέρες", "Θεραπευτής")
    For index = LBound(headers) To UBound(headers)
        If Len(Trim$(CStr(ws.Cells(1, index + 1).Value))) = 0 Then
            ws.Cells(1, index + 1).Value = headers(index)
        ElseIf StrComp(Trim$(CStr(ws.Cells(1, index + 1).Value)), CStr(headers(index)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 2721, , "Το OUTPATIENT_SCHEDULE έχει μη συμβατή δομή στηλών."
        End If
    Next index

    Set EnsureOutpatientScheduleSheet = ws
End Function

Private Function FindPatientRow(ByVal ws As Worksheet, ByVal patientId As String) As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 1).Value)), patientId, vbTextCompare) = 0 Then
            FindPatientRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FindHeader(ByVal ws As Worksheet, ByVal candidates As Variant) As Long
    Dim lastCol As Long
    Dim col As Long
    Dim candidate As Variant
    Dim text As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        text = Trim$(CStr(ws.Cells(1, col).Value))
        For Each candidate In candidates
            If StrComp(text, CStr(candidate), vbTextCompare) = 0 Then
                FindHeader = col
                Exit Function
            End If
        Next candidate
    Next col
End Function

Private Function IsOutpatientValue(ByVal valueText As String) As Boolean
    Dim normalized As String
    normalized = LCase$(Trim$(valueText))
    IsOutpatientValue = (normalized = LCase$("Εξωτερικός") Or normalized = "outpatient")
End Function
'''


def install_outpatient_in_session_writer(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)
