from __future__ import annotations


CLOSE_DAY_MODULE_NAME = "modCloseDayAction"
CLOSE_DAY_MACRO_NAME = "CloseOperationalDay"


CLOSE_DAY_MODULE_CODE = r'''Option Explicit

Private Const SNAPSHOT_SHEET As String = "FINALIZED_DAY_SNAPSHOTS"
Private Const THERAPIST_SHEET As String = "THERAPIST_DAY_HISTORY"
Private Const PATIENT_SHEET As String = "PATIENT_SESSION_HISTORY"
Private Const EVENT_SHEET As String = "DAY_HISTORY_EVENTS"

Public Sub CloseOperationalDay()
    Dim scriptPath As String, payloadPath As String, responsePath As String
    Dim commandLine As String, responseText As String
    Dim answer As VbMsgBoxResult, exitCode As Long

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Ανοίξτε το κανονικά και δοκιμάστε ξανά.", vbExclamation, "Κλείσιμο ημέρας"
        Exit Sub
    End If

    answer = MsgBox( _
        "Να κλείσει οριστικά η επιλεγμένη ημέρα;" & vbCrLf & vbCrLf & _
        "Μετά το κλείσιμο η ημέρα θα περάσει στο ιστορικό.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Κλείσιμο ημέρας")
    If answer <> vbYes Then Exit Sub

    scriptPath = ResolveCloseDayPreparationScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Η λειτουργία Κλεισίματος ημέρας δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
        Exit Sub
    End If

    ThisWorkbook.Save
    payloadPath = Environ$("TEMP") & "\\eka_close_day_payload_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\\eka_close_day_response_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --actor-or-source " & QuoteArg("excel:" & Environ$("USERNAME"))

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)

    If InStr(1, responseText, "already finalized", vbTextCompare) > 0 Then
        MsgBox "Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.", vbInformation, "Κλείσιμο ημέρας"
        GoTo CleanUp
    End If

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then
        MsgBox "Η ημέρα δεν έκλεισε." & vbCrLf & vbCrLf & _
               "Ελέγξτε ότι δεν υπάρχουν εκκρεμείς συνεδρίες και δοκιμάστε ξανά." & vbCrLf & _
               "Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Κλείσιμο ημέρας"
        GoTo CleanUp
    End If

    If Not PersistCloseDayPayload(payloadPath) Then GoTo CleanUp

    MsgBox "Η ημέρα έκλεισε επιτυχώς.", vbInformation, "Κλείσιμο ημέρας"
    ThisWorkbook.Worksheets("MASTER_SCHEDULE").Activate

CleanUp:
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
    Exit Sub

ActionError:
    MsgBox "Η ημέρα δεν έκλεισε. Δοκιμάστε ξανά. Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
    Resume CleanUp
End Sub

Private Function PersistCloseDayPayload(ByVal payloadPath As String) As Boolean
    Dim rawText As String, lines As Variant, parts As Variant
    Dim i As Long, therapistCount As Long, patientCount As Long
    Dim snapshotParts As Variant, eventParts As Variant
    Dim therapistLines() As String, patientLines() As String
    Dim snapshotCreated As Boolean, therapistCreated As Boolean, patientCreated As Boolean, eventCreated As Boolean
    Dim snapshotRow As Long, therapistStartRow As Long, patientStartRow As Long, eventRow As Long
    Dim wsSnapshot As Worksheet, wsTherapist As Worksheet, wsPatient As Worksheet, wsEvent As Worksheet

    On Error GoTo PersistError

    rawText = ReadUtf8Text(payloadPath)
    lines = Split(Replace(rawText, vbCr, ""), vbLf)
    therapistCount = 0
    patientCount = 0

    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo NextLine
        parts = Split(CStr(lines(i)), vbTab)
        Select Case CStr(parts(0))
            Case "SNAPSHOT"
                If UBound(parts) <> 15 Then Err.Raise vbObjectError + 2101, , "Invalid snapshot payload"
                snapshotParts = parts
            Case "THERAPIST"
                If UBound(parts) <> 12 Then Err.Raise vbObjectError + 2102, , "Invalid therapist payload"
                therapistCount = therapistCount + 1
                ReDim Preserve therapistLines(1 To therapistCount)
                therapistLines(therapistCount) = CStr(lines(i))
            Case "PATIENT"
                If UBound(parts) <> 14 Then Err.Raise vbObjectError + 2106, , "Invalid patient payload"
                patientCount = patientCount + 1
                ReDim Preserve patientLines(1 To patientCount)
                patientLines(patientCount) = CStr(lines(i))
            Case "EVENT"
                If UBound(parts) <> 12 Then Err.Raise vbObjectError + 2103, , "Invalid event payload"
                eventParts = parts
            Case Else
                Err.Raise vbObjectError + 2104, , "Unsupported close-day payload row"
        End Select
NextLine:
    Next i

    If IsEmpty(snapshotParts) Or IsEmpty(eventParts) Or therapistCount = 0 Then
        Err.Raise vbObjectError + 2105, , "Incomplete close-day payload"
    End If

    If Not PreflightSnapshot(snapshotParts) Then Exit Function
    If Not PreflightTherapists(therapistLines, therapistCount) Then Exit Function
    If Not PreflightPatients(patientLines, patientCount) Then Exit Function
    If Not PreflightEvent(eventParts) Then Exit Function

    Set wsSnapshot = EnsureHistorySheet(SNAPSHOT_SHEET, SnapshotHeaders(), snapshotCreated)
    Set wsTherapist = EnsureHistorySheet(THERAPIST_SHEET, TherapistHeaders(), therapistCreated)
    Set wsPatient = EnsureHistorySheet(PATIENT_SHEET, PatientHeaders(), patientCreated)
    Set wsEvent = EnsureHistorySheet(EVENT_SHEET, EventHeaders(), eventCreated)

    snapshotRow = NextWriteRow(wsSnapshot)
    therapistStartRow = NextWriteRow(wsTherapist)
    patientStartRow = NextWriteRow(wsPatient)
    eventRow = NextWriteRow(wsEvent)

    WritePayloadFields wsSnapshot, snapshotRow, snapshotParts, 1, 15
    If Not RowMatchesPayload(wsSnapshot, snapshotRow, snapshotParts, 1, 15) Then Err.Raise vbObjectError + 2110, , "Snapshot read-back mismatch"

    For i = 1 To therapistCount
        parts = Split(therapistLines(i), vbTab)
        WritePayloadFields wsTherapist, therapistStartRow + i - 1, parts, 1, 12
        If Not RowMatchesPayload(wsTherapist, therapistStartRow + i - 1, parts, 1, 12) Then Err.Raise vbObjectError + 2111, , "Therapist history read-back mismatch"
    Next i

    For i = 1 To patientCount
        parts = Split(patientLines(i), vbTab)
        WritePayloadFields wsPatient, patientStartRow + i - 1, parts, 1, 14
        If Not RowMatchesPayload(wsPatient, patientStartRow + i - 1, parts, 1, 14) Then Err.Raise vbObjectError + 2113, , "Patient history read-back mismatch"
    Next i

    WritePayloadFields wsEvent, eventRow, eventParts, 1, 12
    If Not RowMatchesPayload(wsEvent, eventRow, eventParts, 1, 12) Then Err.Raise vbObjectError + 2112, , "Event read-back mismatch"

    FinalizeTechnicalSheet wsSnapshot
    FinalizeTechnicalSheet wsTherapist
    FinalizeTechnicalSheet wsPatient
    FinalizeTechnicalSheet wsEvent
    ThisWorkbook.Save
    PersistCloseDayPayload = True
    Exit Function

PersistError:
    On Error Resume Next
    RollbackHistoryWrite wsSnapshot, snapshotCreated, snapshotRow, IIf(snapshotRow > 0, 1, 0)
    RollbackHistoryWrite wsTherapist, therapistCreated, therapistStartRow, therapistCount
    RollbackHistoryWrite wsPatient, patientCreated, patientStartRow, patientCount
    RollbackHistoryWrite wsEvent, eventCreated, eventRow, IIf(eventRow > 0, 1, 0)
    On Error GoTo 0
    MsgBox "Η ημέρα δεν έκλεισε. Δεν αποθηκεύτηκαν αλλαγές. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
End Function

Private Function PreflightSnapshot(ByVal parts As Variant) As Boolean
    Dim ws As Worksheet, row As Long, lastRow As Long
    Set ws = ExistingSheet(SNAPSHOT_SHEET)
    If ws Is Nothing Then PreflightSnapshot = True: Exit Function
    If Not HeadersMatch(ws, SnapshotHeaders()) Then GoTo InvalidSchema
    lastRow = LastDataRow(ws)
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = PayloadValue(parts(1)) Then GoTo Duplicate
        If NormalizedDateText(ws.Cells(row, 2).Value) = PayloadValue(parts(2)) And UCase$(Trim$(CStr(ws.Cells(row, 4).Value))) = "CURRENT_FINAL" Then GoTo Duplicate
    Next row
    PreflightSnapshot = True
    Exit Function
InvalidSchema:
    MsgBox "Το ιστορικό ημέρας δεν είναι διαθέσιμο. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
    Exit Function
Duplicate:
    MsgBox "Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.", vbInformation, "Κλείσιμο ημέρας"
End Function

Private Function PreflightTherapists(ByRef therapistLines() As String, ByVal count As Long) As Boolean
    Dim ws As Worksheet, lastRow As Long, row As Long, i As Long, parts As Variant
    Set ws = ExistingSheet(THERAPIST_SHEET)
    If ws Is Nothing Then PreflightTherapists = True: Exit Function
    If Not HeadersMatch(ws, TherapistHeaders()) Then GoTo InvalidSchema
    lastRow = LastDataRow(ws)
    For i = 1 To count
        parts = Split(therapistLines(i), vbTab)
        For row = 2 To lastRow
            If Trim$(CStr(ws.Cells(row, 1).Value)) = PayloadValue(parts(1)) And Trim$(CStr(ws.Cells(row, 3).Value)) = PayloadValue(parts(3)) Then GoTo Duplicate
        Next row
    Next i
    PreflightTherapists = True
    Exit Function
InvalidSchema:
    MsgBox "Το ιστορικό θεραπευτών δεν είναι διαθέσιμο. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
    Exit Function
Duplicate:
    MsgBox "Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.", vbInformation, "Κλείσιμο ημέρας"
End Function

Private Function PreflightPatients(ByRef patientLines() As String, ByVal count As Long) As Boolean
    Dim ws As Worksheet, lastRow As Long, row As Long, i As Long, parts As Variant
    Set ws = ExistingSheet(PATIENT_SHEET)
    If ws Is Nothing Then PreflightPatients = True: Exit Function
    If Not HeadersMatch(ws, PatientHeaders()) Then GoTo InvalidSchema
    lastRow = LastDataRow(ws)
    For i = 1 To count
        parts = Split(patientLines(i), vbTab)
        For row = 2 To lastRow
            If Trim$(CStr(ws.Cells(row, 1).Value)) = PayloadValue(parts(1)) And Trim$(CStr(ws.Cells(row, 4).Value)) = PayloadValue(parts(4)) Then GoTo Duplicate
        Next row
    Next i
    PreflightPatients = True
    Exit Function
InvalidSchema:
    MsgBox "Το ιστορικό συνεδριών ασθενών δεν είναι διαθέσιμο. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
    Exit Function
Duplicate:
    MsgBox "Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.", vbInformation, "Κλείσιμο ημέρας"
End Function

Private Function PreflightEvent(ByVal parts As Variant) As Boolean
    Dim ws As Worksheet, row As Long, lastRow As Long
    Set ws = ExistingSheet(EVENT_SHEET)
    If ws Is Nothing Then PreflightEvent = True: Exit Function
    If Not HeadersMatch(ws, EventHeaders()) Then GoTo InvalidSchema
    lastRow = LastDataRow(ws)
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = PayloadValue(parts(1)) Then GoTo Duplicate
    Next row
    PreflightEvent = True
    Exit Function
InvalidSchema:
    MsgBox "Το ιστορικό ενεργειών δεν είναι διαθέσιμο. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
    Exit Function
Duplicate:
    MsgBox "Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.", vbInformation, "Κλείσιμο ημέρας"
End Function

Private Function EnsureHistorySheet(ByVal sheetName As String, ByVal headers As Variant, ByRef created As Boolean) As Worksheet
    Dim ws As Worksheet, i As Long
    Set ws = ExistingSheet(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        created = True
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value2 = headers(i)
        Next i
        ws.Rows(1).Font.Bold = True
    Else
        ws.Unprotect
    End If
    Set EnsureHistorySheet = ws
End Function

Private Function ExistingSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set ExistingSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function HeadersMatch(ByVal ws As Worksheet, ByVal headers As Variant) As Boolean
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If Trim$(CStr(ws.Cells(1, i + 1).Value)) <> CStr(headers(i)) Then Exit Function
    Next i
    HeadersMatch = True
End Function

Private Sub WritePayloadFields(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal parts As Variant, ByVal firstPart As Long, ByVal fieldCount As Long)
    Dim i As Long
    For i = 0 To fieldCount - 1
        ws.Cells(targetRow, i + 1).Value2 = PayloadValue(parts(firstPart + i))
    Next i
End Sub

Private Function RowMatchesPayload(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal parts As Variant, ByVal firstPart As Long, ByVal fieldCount As Long) As Boolean
    Dim i As Long
    For i = 0 To fieldCount - 1
        If Not CellMatchesPayload(ws.Cells(targetRow, i + 1).Value, PayloadValue(parts(firstPart + i))) Then Exit Function
    Next i
    RowMatchesPayload = True
End Function

Private Function CellMatchesPayload(ByVal actualValue As Variant, ByVal expectedText As String) As Boolean
    Dim actualText As String
    actualText = Trim$(CStr(actualValue))
    expectedText = Trim$(expectedText)

    If actualText = expectedText Then CellMatchesPayload = True: Exit Function
    If IsNumeric(actualValue) And IsNumeric(expectedText) Then CellMatchesPayload = (CDbl(actualValue) = CDbl(expectedText)): Exit Function
    If IsIsoDateText(expectedText) And IsDate(actualValue) Then CellMatchesPayload = (Format$(CDate(actualValue), "yyyy-mm-dd") = expectedText): Exit Function
    If IsIsoTimeText(expectedText) And IsDate(actualValue) Then CellMatchesPayload = (Format$(CDate(actualValue), "hh:nn:ss") = expectedText): Exit Function
End Function

Private Function IsIsoDateText(ByVal value As String) As Boolean
    If Len(value) <> 10 Then Exit Function
    If Mid$(value, 5, 1) <> "-" Or Mid$(value, 8, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(value, 4)) Or Not IsNumeric(Mid$(value, 6, 2)) Or Not IsNumeric(Right$(value, 2)) Then Exit Function
    IsIsoDateText = True
End Function

Private Function IsIsoTimeText(ByVal value As String) As Boolean
    If Len(value) <> 8 Then Exit Function
    If Mid$(value, 3, 1) <> ":" Or Mid$(value, 6, 1) <> ":" Then Exit Function
    If Not IsNumeric(Left$(value, 2)) Or Not IsNumeric(Mid$(value, 4, 2)) Or Not IsNumeric(Right$(value, 2)) Then Exit Function
    IsIsoTimeText = True
End Function

Private Function PayloadValue(ByVal raw As String) As String
    Dim i As Long, ch As String, result As String, escaped As Boolean
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If escaped Then
            Select Case ch
                Case "t": result = result & vbTab
                Case "r": result = result & vbCr
                Case "n": result = result & vbLf
                Case "\": result = result & "\"
                Case Else: result = result & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        Else
            result = result & ch
        End If
    Next i
    If escaped Then result = result & "\"
    PayloadValue = result
End Function

Private Function NormalizedDateText(ByVal value As Variant) As String
    If IsDate(value) Then NormalizedDateText = Format$(CDate(value), "yyyy-mm-dd") Else NormalizedDateText = Trim$(CStr(value))
End Function

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LastDataRow < 1 Then LastDataRow = 1
End Function

Private Function NextWriteRow(ByVal ws As Worksheet) As Long
    NextWriteRow = LastDataRow(ws) + 1
    If NextWriteRow < 2 Then NextWriteRow = 2
End Function

Private Sub FinalizeTechnicalSheet(ByVal ws As Worksheet)
    ws.Visible = xlSheetVeryHidden
    ws.Protect
End Sub

Private Sub RollbackHistoryWrite(ByVal ws As Worksheet, ByVal created As Boolean, ByVal startRow As Long, ByVal rowCount As Long)
    Dim oldAlerts As Boolean
    If ws Is Nothing Then Exit Sub
    If created Then
        oldAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = oldAlerts
    ElseIf startRow >= 2 And rowCount > 0 Then
        ws.Unprotect
        ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + rowCount - 1, ws.UsedRange.Columns.Count)).ClearContents
        FinalizeTechnicalSheet ws
    End If
End Sub

Private Function SnapshotHeaders() As Variant
    SnapshotHeaders = Array("SnapshotId", "ServiceDate", "Version", "Status", "FinalizedAt", "SupersedesSnapshotId", "UnresolvedSessions", "PlannedSessionCount", "FinalSessionCount", "PatientAbsenceCount", "CancellationCount", "NoShowCount", "ReplacementCount", "TreatmentSuspensionCount", "ApprovedExceptionCount")
End Function

Private Function TherapistHeaders() As Variant
    TherapistHeaders = Array("SnapshotId", "ServiceDate", "TherapistId", "CapacityTimeslots", "PlannedOccupiedTimeslots", "PlannedSessions", "FinalOccupiedTimeslots", "FinalActiveSessions", "ReplacementsReceived", "ReleasedByCancellationOrNoShow", "TreatmentSuspensionsAffectingLoad", "ApprovedSharedOrOverCapacityEvents")
End Function

Private Function PatientHeaders() As Variant
    PatientHeaders = Array("SnapshotId", "ServiceDate", "SnapshotVersion", "SessionId", "PatientId", "Status", "OriginalTherapistId", "OriginalTime", "EffectiveTherapistId", "EffectiveTime", "ReplacementId", "CancellationId", "CancellationKind", "Reason")
End Function

Private Function EventHeaders() As Variant
    EventHeaders = Array("EventId", "ServiceDate", "EventType", "RecordedAt", "ActorOrSource", "TherapistId", "PatientId", "SessionId", "ReasonCode", "RelatedEventId", "RelatedSnapshotId", "Summary")
End Function

Private Function ResolveCloseDayPreparationScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\\Python\\tools\\prepare_close_day_in_place.py"
    If Dir$(candidate) <> "" Then ResolveCloseDayPreparationScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\prepare_close_day_in_place.py"
    If Dir$(candidate) <> "" Then ResolveCloseDayPreparationScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
End Function

Private Function ReadUtf8Text(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8Text = stream.ReadText
    stream.Close
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
'''

# VBA string literals do not use backslash escaping. Normalize the Python raw-string
# source before installation so paths and one-character payload escape parsing use
# a single Windows separator/backslash at runtime.
CLOSE_DAY_MODULE_CODE = CLOSE_DAY_MODULE_CODE.replace("\\\\", "\\")


def install_close_day_action(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(CLOSE_DAY_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = CLOSE_DAY_MODULE_NAME
    module.CodeModule.AddFromString(CLOSE_DAY_MODULE_CODE)
