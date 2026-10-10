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
    If Len(timeText) = 0 Or Not IsDate(timeText) Then Err.Raise vbObjectError + 2712, , "Η ώρα δεν είναι έγκυρη."
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
    ws.Cells(targetRow, 4).Value = TimeValue(timeText)
    ws.Cells(targetRow, 4).NumberFormat = "hh:mm"
    ws.Cells(targetRow, 5).Value = dayPattern
    ws.Cells(targetRow, 6).Value = therapistName

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
