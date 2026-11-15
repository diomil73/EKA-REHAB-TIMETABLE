from __future__ import annotations


DOUBLE_BOOKING_PROJECTION_MODULE_NAME = "modDoubleBookingDailyProjection"


MODULE_CODE = r'''Option Explicit

Public Function ProjectDoubleBookingToTherapistDaily( _
    ByVal patientName As String, _
    ByVal therapistName As String, _
    ByVal replacementTime As String, _
    ByVal targetSessionId As String) As Boolean

    Dim ws As Worksheet, targetCell As Range
    Dim existingText As String

    On Error GoTo ProjectionError

    Set ws = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    Set targetCell = ResolveTherapistDailyCell(ws, therapistName, replacementTime)
    If targetCell Is Nothing Then
        RollbackValidatedFact targetSessionId
        MsgBox "Δεν βρέθηκε μοναδική θέση στο ημερήσιο πρόγραμμα για " & _
               therapistName & " στις " & replacementTime & ". Η διπλοκράτηση αναιρέθηκε.", _
               vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    existingText = CStr(targetCell.Value)
    If Not TextHasExactLine(existingText, patientName) Then
        If Len(Trim$(existingText)) > 0 Then
            targetCell.Value = existingText & vbLf & patientName
        Else
            targetCell.Value = patientName
        End If
    End If

    targetCell.WrapText = True
    If targetCell.Font.Size < 10 Then targetCell.Font.Size = 10

    If Not TextHasExactLine(CStr(targetCell.Value), patientName) Then
        RollbackValidatedFact targetSessionId
        MsgBox "Η διπλοκράτηση δεν εμφανίστηκε σωστά στο ημερήσιο πρόγραμμα και αναιρέθηκε.", _
               vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    ThisWorkbook.Save
    ProjectDoubleBookingToTherapistDaily = True
    Exit Function

ProjectionError:
    On Error Resume Next
    RollbackValidatedFact targetSessionId
    On Error GoTo 0
    MsgBox "Δεν ήταν δυνατή η ενημέρωση του ημερήσιου προγράμματος: " & Err.Description, _
           vbCritical, "Διπλοκράτηση"
End Function

Private Function ResolveTherapistDailyCell( _
    ByVal ws As Worksheet, _
    ByVal therapistName As String, _
    ByVal replacementTime As String) As Range

    Dim headerRows As Variant, firstRows As Variant, lastRows As Variant
    Dim blockIndex As Long, col As Long, row As Long
    Dim providerCol As Long, matchCount As Long
    Dim candidate As Range

    headerRows = Array(1, 11)
    firstRows = Array(2, 12)
    lastRows = Array(8, 18)

    For blockIndex = 0 To 1
        providerCol = 0
        For col = 2 To 10
            If NormalizeName(CStr(ws.Cells(headerRows(blockIndex), col).Value)) = _
               NormalizeName(therapistName) Then
                providerCol = col
                Exit For
            End If
        Next col

        If providerCol > 0 Then
            For row = firstRows(blockIndex) To lastRows(blockIndex)
                If NormalizedTimeText(ws.Cells(row, 1).Value) = replacementTime Then
                    matchCount = matchCount + 1
                    Set candidate = ws.Cells(row, providerCol)
                End If
            Next row
        End If
    Next blockIndex

    If matchCount = 1 Then Set ResolveTherapistDailyCell = candidate
End Function

Private Function NormalizedTimeText(ByVal value As Variant) As String
    If IsDate(value) Then
        NormalizedTimeText = Format$(CDate(value), "hh:mm")
    ElseIf IsNumeric(value) Then
        NormalizedTimeText = Format$(CDbl(value), "hh:mm")
    Else
        NormalizedTimeText = Left$(Trim$(CStr(value)), 5)
    End If
End Function

Private Function NormalizeName(ByVal value As String) As String
    Dim text As String
    text = UCase$(Trim$(value))
    text = Replace(text, "Ά", "Α")
    text = Replace(text, "Έ", "Ε")
    text = Replace(text, "Ή", "Η")
    text = Replace(text, "Ί", "Ι")
    text = Replace(text, "Ϊ", "Ι")
    text = Replace(text, "Ό", "Ο")
    text = Replace(text, "Ύ", "Υ")
    text = Replace(text, "Ϋ", "Υ")
    text = Replace(text, "Ώ", "Ω")
    NormalizeName = text
End Function

Private Function TextHasExactLine(ByVal text As String, ByVal wanted As String) As Boolean
    Dim lines As Variant, i As Long
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If NormalizeName(CStr(lines(i))) = NormalizeName(wanted) Then
            TextHasExactLine = True
            Exit Function
        End If
    Next i
End Function

Private Sub RollbackValidatedFact(ByVal targetSessionId As String)
    Const FACT_SHEET As String = "DAILY_REPLACEMENT_FACTS"
    Dim ws As Worksheet
    Dim row As Long, lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For row = lastRow To 2 Step -1
        If Trim$(CStr(ws.Cells(row, 2).Value)) = targetSessionId And _
           UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = "ACTIVE" And _
           Trim$(CStr(ws.Cells(row, 8).Value)) = "CONFIRMED DOUBLE BOOKING EXCEPTION" Then
            ws.Rows(row).ClearContents
            Exit For
        End If
    Next row
    ThisWorkbook.Save
End Sub
'''


def install_double_booking_daily_projection(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(DOUBLE_BOOKING_PROJECTION_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = DOUBLE_BOOKING_PROJECTION_MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)
