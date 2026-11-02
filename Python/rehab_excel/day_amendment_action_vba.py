from __future__ import annotations

DAY_AMENDMENT_MODULE_NAME = "modDayAmendmentAction"
DAY_AMENDMENT_MACRO_NAME = "StartClosedDayAmendment"
DAY_AMENDMENT_BUTTON_NAME = "ekaToolbar_DayAmendment"

DAY_AMENDMENT_MODULE_CODE = r'''Option Explicit

Private Const SNAPSHOT_SHEET As String = "FINALIZED_DAY_SNAPSHOTS"
Private Const THERAPIST_SHEET As String = "THERAPIST_DAY_HISTORY"
Private Const EVENT_SHEET As String = "DAY_HISTORY_EVENTS"

Public Sub StartClosedDayAmendment()
    Dim prepareScript As String
    Dim payloadPath As String, responsePath As String
    Dim commandLine As String, responseText As String
    Dim actorOrSource As String
    Dim answer As VbMsgBoxResult, exitCode As Long

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Ανοίξτε το κανονικά και δοκιμάστε ξανά.", vbExclamation, "Αποθήκευση ημέρας"
        Exit Sub
    End If

    answer = MsgBox( _
        "Να αποθηκευτεί η κλεισμένη ημέρα με τις τρέχουσες αλλαγές;" & vbCrLf & vbCrLf & _
        "Η προηγούμενη καταχώρηση θα παραμείνει στο ιστορικό.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Αποθήκευση ημέρας")
    If answer <> vbYes Then Exit Sub

    prepareScript = ResolveRepoTool("prepare_day_amendment_in_session.py")
    If Len(prepareScript) = 0 Then
        MsgBox "Η λειτουργία Αποθήκευσης ημέρας δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Αποθήκευση ημέρας"
        Exit Sub
    End If

    ThisWorkbook.Save

    actorOrSource = "excel:" & Environ$("USERNAME")
    payloadPath = Environ$("TEMP") & "\\eka_day_save_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".tsv"
    responsePath = Environ$("TEMP") & "\\eka_day_save_prepare_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(prepareScript) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --actor-or-source " & QuoteArg(actorOrSource)

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8File(responsePath)

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then
        MsgBox "Η ημέρα δεν αποθηκεύτηκε." & vbCrLf & vbCrLf & _
               "Ελέγξτε ότι η ημέρα είναι κλεισμένη και ότι δεν υπάρχουν εκκρεμείς συνεδρίες." & vbCrLf & _
               "Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Αποθήκευση ημέρας"
        GoTo CleanUp
    End If

    If Not PersistDaySavePayload(payloadPath) Then GoTo CleanUp

    MsgBox "Η ημέρα αποθηκεύτηκε επιτυχώς.", vbInformation, "Αποθήκευση ημέρας"

CleanUp:
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
    Exit Sub

ActionError:
    MsgBox "Η ημέρα δεν αποθηκεύτηκε. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", vbCritical, "Αποθήκευση ημέρας"
    Resume CleanUp
End Sub

Private Function PersistDaySavePayload(ByVal payloadPath As String) As Boolean
    Dim rawText As String, lines As Variant, parts As Variant
    Dim supersedeParts As Variant, snapshotParts As Variant
    Dim therapistLines() As String, eventLines() As String
    Dim therapistCount As Long, eventCount As Long, i As Long
    Dim wsSnapshot As Worksheet, wsTherapist As Worksheet, wsEvent As Worksheet
    Dim oldSnapshotRow As Long, newSnapshotRow As Long
    Dim therapistStartRow As Long, eventStartRow As Long
    Dim oldStatus As String

    On Error GoTo PersistError

    rawText = Replace(ReadUtf8File(payloadPath), ChrW(&HFEFF), "")
    lines = Split(Replace(rawText, vbCr, ""), vbLf)

    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo NextLine
        parts = Split(CStr(lines(i)), vbTab)
        Select Case CStr(parts(0))
            Case "SUPERSEDE"
                If UBound(parts) <> 15 Then Err.Raise vbObjectError + 7401, , "Invalid supersede payload"
                supersedeParts = parts
            Case "SNAPSHOT"
                If UBound(parts) <> 15 Then Err.Raise vbObjectError + 7402, , "Invalid snapshot payload"
                snapshotParts = parts
            Case "THERAPIST"
                If UBound(parts) <> 12 Then Err.Raise vbObjectError + 7403, , "Invalid therapist payload"
                therapistCount = therapistCount + 1
                ReDim Preserve therapistLines(1 To therapistCount)
                therapistLines(therapistCount) = CStr(lines(i))
            Case "EVENT"
                If UBound(parts) <> 12 Then Err.Raise vbObjectError + 7404, , "Invalid event payload"
                eventCount = eventCount + 1
                ReDim Preserve eventLines(1 To eventCount)
                eventLines(eventCount) = CStr(lines(i))
            Case Else
                Err.Raise vbObjectError + 7405, , "Unsupported day-save payload row"
        End Select
NextLine:
    Next i

    If IsEmpty(supersedeParts) Or IsEmpty(snapshotParts) Or therapistCount = 0 Or eventCount <> 2 Then
        Err.Raise vbObjectError + 7406, , "Incomplete day-save payload"
    End If

    Set wsSnapshot = ExistingSheet(SNAPSHOT_SHEET)
    Set wsTherapist = ExistingSheet(THERAPIST_SHEET)
    Set wsEvent = ExistingSheet(EVENT_SHEET)
    If wsSnapshot Is Nothing Or wsTherapist Is Nothing Or wsEvent Is Nothing Then
        Err.Raise vbObjectError + 7407, , "Day-history storage is missing"
    End If

    wsSnapshot.Unprotect
    wsTherapist.Unprotect
    wsEvent.Unprotect

    oldSnapshotRow = FindRowById(wsSnapshot, PayloadValue(supersedeParts(1)))
    If oldSnapshotRow = 0 Then Err.Raise vbObjectError + 7408, , "Previous snapshot was not found"
    If UCase$(Trim$(CStr(wsSnapshot.Cells(oldSnapshotRow, 4).Value))) <> "CURRENT_FINAL" Then
        Err.Raise vbObjectError + 7409, , "Previous snapshot is not CURRENT_FINAL"
    End If
    If FindRowById(wsSnapshot, PayloadValue(snapshotParts(1))) <> 0 Then
        Err.Raise vbObjectError + 7410, , "Replacement snapshot already exists"
    End If

    For i = 1 To therapistCount
        parts = Split(therapistLines(i), vbTab)
        If HistoryPairExists(wsTherapist, PayloadValue(parts(1)), PayloadValue(parts(3))) Then
            Err.Raise vbObjectError + 7411, , "Replacement therapist history already exists"
        End If
    Next i
    For i = 1 To eventCount
        parts = Split(eventLines(i), vbTab)
        If FindRowById(wsEvent, PayloadValue(parts(1))) <> 0 Then
            Err.Raise vbObjectError + 7412, , "Day-save event already exists"
        End If
    Next i

    oldStatus = CStr(wsSnapshot.Cells(oldSnapshotRow, 4).Value)
    newSnapshotRow = NextWriteRow(wsSnapshot)
    therapistStartRow = NextWriteRow(wsTherapist)
    eventStartRow = NextWriteRow(wsEvent)

    ' The amendment core changes only Status on the prior snapshot. Native Excel
    ' may store ServiceDate as a typed date and Version as a numeric Double, so
    ' compare only the intended fields with numeric normalization for Version.
    wsSnapshot.Cells(oldSnapshotRow, 4).Value2 = PayloadValue(supersedeParts(4))
    If Not SupersededSnapshotMatches(wsSnapshot, oldSnapshotRow, supersedeParts) Then
        Err.Raise vbObjectError + 7420, , "Superseded snapshot read-back mismatch"
    End If

    WritePayloadFields wsSnapshot, newSnapshotRow, snapshotParts, 1, 15
    If Not RowMatchesPayload(wsSnapshot, newSnapshotRow, snapshotParts, 1, 15) Then
        Err.Raise vbObjectError + 7421, , "Replacement snapshot read-back mismatch"
    End If

    For i = 1 To therapistCount
        parts = Split(therapistLines(i), vbTab)
        WritePayloadFields wsTherapist, therapistStartRow + i - 1, parts, 1, 12
        If Not RowMatchesPayload(wsTherapist, therapistStartRow + i - 1, parts, 1, 12) Then
            Err.Raise vbObjectError + 7422, , "Therapist history read-back mismatch"
        End If
    Next i

    For i = 1 To eventCount
        parts = Split(eventLines(i), vbTab)
        WritePayloadFields wsEvent, eventStartRow + i - 1, parts, 1, 12
        If Not RowMatchesPayload(wsEvent, eventStartRow + i - 1, parts, 1, 12) Then
            Err.Raise vbObjectError + 7423, , "Day-save event read-back mismatch"
        End If
    Next i

    FinalizeTechnicalSheet wsSnapshot
    FinalizeTechnicalSheet wsTherapist
    FinalizeTechnicalSheet wsEvent
    ThisWorkbook.Save
    PersistDaySavePayload = True
    Exit Function

PersistError:
    On Error Resume Next
    If Not wsSnapshot Is Nothing Then
        If oldSnapshotRow > 0 Then wsSnapshot.Cells(oldSnapshotRow, 4).Value2 = oldStatus
        If newSnapshotRow > 0 Then wsSnapshot.Rows(newSnapshotRow).ClearContents
    End If
    If Not wsTherapist Is Nothing And therapistStartRow > 0 And therapistCount > 0 Then
        wsTherapist.Rows(CStr(therapistStartRow) & ":" & CStr(therapistStartRow + therapistCount - 1)).ClearContents
    End If
    If Not wsEvent Is Nothing And eventStartRow > 0 And eventCount > 0 Then
        wsEvent.Rows(CStr(eventStartRow) & ":" & CStr(eventStartRow + eventCount - 1)).ClearContents
    End If
    FinalizeTechnicalSheet wsSnapshot
    FinalizeTechnicalSheet wsTherapist
    FinalizeTechnicalSheet wsEvent
    ThisWorkbook.Save
    On Error GoTo 0
    MsgBox "Η ημέρα δεν αποθηκεύτηκε. Δεν γράφτηκε νέα έκδοση στο ιστορικό.", vbCritical, "Αποθήκευση ημέρας"
End Function

Private Function ExistingSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set ExistingSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindRowById(ByVal ws As Worksheet, ByVal recordId As String) As Long
    Dim lastRow As Long, row As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = recordId Then
            FindRowById = row
            Exit Function
        End If
    Next row
End Function

Private Function HistoryPairExists(ByVal ws As Worksheet, ByVal snapshotId As String, ByVal therapistId As String) As Boolean
    Dim lastRow As Long, row As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = snapshotId And Trim$(CStr(ws.Cells(row, 3).Value)) = therapistId Then
            HistoryPairExists = True
            Exit Function
        End If
    Next row
End Function

Private Function NextWriteRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    NextWriteRow = lastRow + 1
    If NextWriteRow < 2 Then NextWriteRow = 2
End Function

Private Sub WritePayloadFields(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal parts As Variant, ByVal firstPart As Long, ByVal fieldCount As Long)
    Dim i As Long
    For i = 0 To fieldCount - 1
        ws.Cells(targetRow, i + 1).Value2 = PayloadValue(parts(firstPart + i))
    Next i
End Sub

Private Function SupersededSnapshotMatches(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal parts As Variant) As Boolean
    If Trim$(CStr(ws.Cells(targetRow, 1).Value)) <> PayloadValue(parts(1)) Then Exit Function
    If Not IsNumeric(ws.Cells(targetRow, 3).Value) Or Not IsNumeric(PayloadValue(parts(3))) Then Exit Function
    If CLng(ws.Cells(targetRow, 3).Value) <> CLng(PayloadValue(parts(3))) Then Exit Function
    If UCase$(Trim$(CStr(ws.Cells(targetRow, 4).Value))) <> UCase$(PayloadValue(parts(4))) Then Exit Function
    SupersededSnapshotMatches = True
End Function

Private Function RowMatchesPayload(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal parts As Variant, ByVal firstPart As Long, ByVal fieldCount As Long) As Boolean
    Dim i As Long
    For i = 0 To fieldCount - 1
        If Trim$(CStr(ws.Cells(targetRow, i + 1).Value)) <> PayloadValue(parts(firstPart + i)) Then Exit Function
    Next i
    RowMatchesPayload = True
End Function

Private Function PayloadValue(ByVal raw As Variant) As String
    PayloadValue = Replace(CStr(raw), ChrW(&HFEFF), "")
End Function

Private Sub FinalizeTechnicalSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Visible = 2
    ws.Protect
    On Error GoTo 0
End Sub

Private Function ResolveRepoTool(ByVal fileName As String) As String
    Dim candidate As String, rootPath As String
    rootPath = ThisWorkbook.Path

    candidate = rootPath & "\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveRepoTool = candidate: Exit Function

    candidate = rootPath & "\\..\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveRepoTool = candidate: Exit Function

    candidate = rootPath & "\\..\\..\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveRepoTool = candidate: Exit Function
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function ReadUtf8File(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadUtf8File = stream.ReadText
    stream.Close
End Function
'''


def _last_used_row(ws) -> int:
    try:
        used = ws.UsedRange
        return max(1, int(used.Row) + int(used.Rows.Count) - 1)
    except Exception:
        return 1


def _delete_button_if_present(ws) -> None:
    try:
        ws.Shapes(DAY_AMENDMENT_BUTTON_NAME).Delete()
    except Exception:
        pass


def install_day_amendment_action(vbproject, workbook) -> None:
    try:
        component = vbproject.VBComponents(DAY_AMENDMENT_MODULE_NAME)
        vbproject.VBComponents.Remove(component)
    except Exception:
        pass

    module = vbproject.VBComponents.Add(1)
    module.Name = DAY_AMENDMENT_MODULE_NAME
    module.CodeModule.AddFromString(DAY_AMENDMENT_MODULE_CODE)

    # This installer owns only the day-save action on DAILY_INPUT.
    # MASTER and the wider application shell are intentionally untouched.
    ws = workbook.Worksheets("DAILY_INPUT")
    _delete_button_if_present(ws)

    last_row = _last_used_row(ws)
    anchor = ws.Range(f"A{last_row + 2}")
    left = float(anchor.Left) + 292
    top = float(anchor.Top) + 2
    shape = ws.Shapes.AddShape(5, left, top, 112, 24)
    shape.Name = DAY_AMENDMENT_BUTTON_NAME
    shape.OnAction = DAY_AMENDMENT_MACRO_NAME
    shape.AlternativeText = "Αποθήκευση αλλαγών σε ήδη κλεισμένη ημέρα με διατήρηση της προηγούμενης καταχώρησης στο ιστορικό."
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.Line.Weight = 1
    shape.TextFrame2.TextRange.Text = "Αποθήκευση ημέρας"
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    try:
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    except Exception:
        pass
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3