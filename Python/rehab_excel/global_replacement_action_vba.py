from __future__ import annotations


GLOBAL_REPLACEMENT_MODULE_NAME = "modGlobalReplacementAction"
GLOBAL_REPLACEMENT_MACRO_NAME = "ResolveDailyReplacementsInPlace"


GLOBAL_REPLACEMENT_MODULE_CODE = r'''Option Explicit

Private Const FACT_SHEET As String = "DAILY_REPLACEMENT_FACTS"

Public Function ResolveDailyReplacementsInPlace() As Boolean
    Dim scriptPath As String, payloadPath As String, responsePath As String
    Dim commandLine As String, responseText As String, mode As String
    Dim exitCode As Long, beforeReview As Long, afterReview As Long

    On Error GoTo ActionError

    scriptPath = ResolveReplacementPreparationScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Η λειτουργία αναπλήρωσης δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", _
               vbCritical, "Αναπληρώσεις"
        Exit Function
    End If

    ThisWorkbook.Save
    payloadPath = Environ$("TEMP") & "\eka_replacement_payload_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\eka_replacement_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(responsePath)

    Application.StatusBar = "Έλεγχος αναπληρώσεων..."
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    Application.StatusBar = False
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)

    If exitCode <> 0 Or _
       InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or _
       Dir$(payloadPath) = "" Then
        MsgBox "Οι αναπληρώσεις δεν εφαρμόστηκαν." & vbCrLf & vbCrLf & _
               "Ελέγξτε τις καταχωρίσεις και δοκιμάστε ξανά." & vbCrLf & _
               "Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Αναπληρώσεις"
        GoTo CleanUp
    End If

    mode = JsonStringValue(responseText, "mode")
    If mode = "already_resolved" Then
        ResolveDailyReplacementsInPlace = True
        GoTo CleanUp
    End If

    If Not PersistReplacementPayload(payloadPath, mode = "resolved") Then GoTo CleanUp

    If mode = "review_required" Then
        beforeReview = ActiveReplacementCount()
        On Error GoTo ReviewUnavailable
        Application.Run "ShowDoubleBookingReview"
        On Error GoTo ActionError
        afterReview = ActiveReplacementCount()
        If afterReview > beforeReview Then
            Application.Run "ApplyDailyInputPreview"
        Else
            MsgBox "Υπάρχουν συνεδρίες που χρειάζονται απόφαση πριν ολοκληρωθεί η ημέρα.", _
                   vbInformation, "Αναπληρώσεις"
        End If
        GoTo CleanUp
    End If

    ResolveDailyReplacementsInPlace = True
    GoTo CleanUp

ReviewUnavailable:
    WriteReplacementDiagnostic Err.Number, Err.Description, "open guided review", ""
    MsgBox "Χρειάζεται επιλογή για μία ή περισσότερες συνεδρίες, αλλά η οθόνη επιλογής δεν είναι διαθέσιμη." & vbCrLf & _
           "Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Αναπληρώσεις"
    Resume CleanUp

ActionError:
    Application.StatusBar = False
    WriteReplacementDiagnostic Err.Number, Err.Description, "replacement action", ""
    MsgBox "Οι αναπληρώσεις δεν ολοκληρώθηκαν. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", _
           vbCritical, "Αναπληρώσεις"

CleanUp:
    Application.StatusBar = False
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
End Function

Private Function PersistReplacementPayload(ByVal payloadPath As String, ByVal applyFinalPatches As Boolean) As Boolean
    Dim rawText As String, lines As Variant, parts As Variant
    Dim replacementLines As Collection, suspensionLines As Collection, patchLines As Collection
    Dim item As Variant, i As Long, patchPath As String
    Dim factCreated As Boolean, factStartRow As Long, factRows As Long
    Dim suspensionRows As Collection
    Dim rollbackPath As String, phase As String, target As String
    Dim previousEvents As Boolean, previousScreenUpdating As Boolean

    On Error GoTo PersistError
    phase = "read payload"
    rawText = ReadUtf8Text(payloadPath)
    lines = Split(Replace(rawText, vbCr, ""), vbLf)
    Set replacementLines = New Collection
    Set suspensionLines = New Collection
    Set patchLines = New Collection
    Set suspensionRows = New Collection

    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo NextLine
        parts = Split(CStr(lines(i)), vbTab)
        Select Case CStr(parts(0))
            Case "REPLACEMENT"
                If UBound(parts) <> 10 Then Err.Raise vbObjectError + 2501, , "Invalid replacement payload"
                replacementLines.Add CStr(lines(i))
            Case "SUSPENSION"
                If UBound(parts) <> 6 Then Err.Raise vbObjectError + 2502, , "Invalid suspension payload"
                suspensionLines.Add CStr(lines(i))
            Case "PATCH"
                If UBound(parts) <> 12 Then Err.Raise vbObjectError + 2503, , "Invalid patch payload"
                patchLines.Add CStr(lines(i))
            Case Else
                Err.Raise vbObjectError + 2504, , "Unsupported replacement payload row"
        End Select
NextLine:
    Next i

    phase = "create rollback"
    rollbackPath = Environ$("TEMP") & "\eka_replacement_rollback_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".xlsm"
    ThisWorkbook.SaveCopyAs rollbackPath

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    phase = "persist replacement facts"
    If replacementLines.Count > 0 Then
        If Not PersistReplacementFacts(replacementLines, factCreated, factStartRow, factRows) Then _
            Err.Raise vbObjectError + 2510, , "Replacement fact write failed"
    End If

    phase = "persist suspensions"
    If suspensionLines.Count > 0 Then
        If Not PersistSuspensions(suspensionLines, suspensionRows) Then _
            Err.Raise vbObjectError + 2511, , "Suspension write failed"
    End If

    If applyFinalPatches Then
        If patchLines.Count = 0 Then Err.Raise vbObjectError + 2512, , "Resolved payload has no projection patches"
        phase = "apply operational projections"
        patchPath = WritePatchPayload(patchLines)
        Application.EnableEvents = previousEvents
        Application.ScreenUpdating = previousScreenUpdating
        If Not ApplyValidatedDailyInputPayload(patchPath) Then GoTo PersistRollback
        previousEvents = Application.EnableEvents
        previousScreenUpdating = Application.ScreenUpdating
        Application.EnableEvents = False
        Application.ScreenUpdating = False
    End If

    phase = "save"
    ThisWorkbook.Save
    PersistReplacementPayload = True
    GoTo PersistExit

PersistRollback:
    phase = "rollback after projection failure"
    GoTo PersistError

PersistError:
    target = ""
    WriteReplacementDiagnostic Err.Number, Err.Description, phase, target
    On Error Resume Next
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then
        RestoreReplacementTransaction rollbackPath, replacementLines, suspensionRows, patchLines, factCreated, factStartRow, factRows
        ThisWorkbook.Save
    End If
    On Error GoTo 0
    MsgBox "Οι αναπληρώσεις δεν εφαρμόστηκαν και οι αλλαγές αναιρέθηκαν." & vbCrLf & _
           "Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", vbCritical, "Αναπληρώσεις"

PersistExit:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    On Error Resume Next
    If Len(patchPath) > 0 And Dir$(patchPath) <> "" Then Kill patchPath
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then Kill rollbackPath
    On Error GoTo 0
End Function

Private Function PersistReplacementFacts( _
    ByVal replacementLines As Collection, _
    ByRef created As Boolean, _
    ByRef startRow As Long, _
    ByRef rowCount As Long) As Boolean

    Dim ws As Worksheet, lastRow As Long, row As Long, i As Long
    Dim parts As Variant, targetId As String, replacementId As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = FACT_SHEET
        WriteFactHeaders ws
        created = True
    ElseIf Not FactHeadersAreValid(ws) Then
        GoTo Failed
    End If
    ws.Unprotect
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    startRow = IIf(lastRow < 2, 2, lastRow + 1)

    For i = 1 To replacementLines.Count
        parts = Split(CStr(replacementLines(i)), vbTab)
        replacementId = PayloadValue(parts(1))
        targetId = PayloadValue(parts(2))
        For row = 2 To lastRow
            If Trim$(CStr(ws.Cells(row, 1).Value)) = replacementId Then GoTo Failed
            If Trim$(CStr(ws.Cells(row, 2).Value)) = targetId And _
               UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = "ACTIVE" Then GoTo Failed
        Next row
    Next i

    For i = 1 To replacementLines.Count
        parts = Split(CStr(replacementLines(i)), vbTab)
        row = startRow + i - 1
        ws.Cells(row, 1).Value2 = PayloadValue(parts(1))
        ws.Cells(row, 2).Value2 = PayloadValue(parts(2))
        ws.Cells(row, 3).Value2 = PayloadValue(parts(3))
        ws.Cells(row, 4).Value2 = PayloadValue(parts(4))
        ws.Cells(row, 5).Value2 = PayloadValue(parts(5))
        ws.Cells(row, 6).NumberFormat = "@"
        ws.Cells(row, 6).Value2 = PayloadValue(parts(6))
        ws.Cells(row, 7).NumberFormat = "@"
        ws.Cells(row, 7).Value2 = PayloadValue(parts(7))
        ws.Cells(row, 8).Value2 = PayloadValue(parts(8))
        ws.Cells(row, 9).Value2 = PayloadValue(parts(9))
        ws.Cells(row, 10).Value2 = PayloadValue(parts(10))
        If Trim$(CStr(ws.Cells(row, 1).Value)) <> PayloadValue(parts(1)) Or _
           Trim$(CStr(ws.Cells(row, 2).Value)) <> PayloadValue(parts(2)) Or _
           UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) <> "ACTIVE" Then GoTo Failed
    Next i
    rowCount = replacementLines.Count
    ws.Visible = xlSheetVeryHidden
    ws.Protect
    PersistReplacementFacts = True
    Exit Function
Failed:
    PersistReplacementFacts = False
End Function

Private Function PersistSuspensions(ByVal suspensionLines As Collection, ByRef writtenRows As Collection) As Boolean
    Dim ws As Worksheet, headerRow As Long, firstRow As Long, lastRow As Long
    Dim emptyRows As Collection, row As Long, i As Long, targetRow As Long, parts As Variant

    Set ws = ThisWorkbook.Worksheets("DAILY_INPUT")
    headerRow = FindPatientHeaderRow(ws)
    If headerRow = 0 Then Exit Function
    firstRow = headerRow + 1
    lastRow = firstRow + 19
    Set emptyRows = New Collection
    For row = firstRow To lastRow
        If Len(Trim$(CStr(ws.Cells(row, 1).Value))) = 0 Then emptyRows.Add row
    Next row
    If emptyRows.Count < suspensionLines.Count Then Exit Function

    For i = 1 To suspensionLines.Count
        parts = Split(CStr(suspensionLines(i)), vbTab)
        targetRow = CLng(emptyRows(i))
        writtenRows.Add targetRow
        ws.Cells(targetRow, 1).Value2 = PayloadValue(parts(3))
        ws.Cells(targetRow, 2).Value2 = "ΟΧΙ"
        ws.Cells(targetRow, 3).NumberFormat = "@"
        ws.Cells(targetRow, 3).Value2 = PayloadValue(parts(4))
        ws.Cells(targetRow, 4).Value2 = PayloadValue(parts(5))
        ws.Cells(targetRow, 5).Value2 = PayloadValue(parts(6))
        If Trim$(CStr(ws.Cells(targetRow, 4).Value)) <> "Αναστολή θεραπείας" Then Exit Function
    Next i
    PersistSuspensions = True
End Function

Private Function FindPatientHeaderRow(ByVal ws As Worksheet) As Long
    Dim row As Long
    For row = 1 To 100
        If UCase$(Trim$(CStr(ws.Cells(row, 1).Value))) = "ΑΣΘΕΝΗΣ" And _
           UCase$(Trim$(CStr(ws.Cells(row, 2).Value))) = "ΟΛΗ ΗΜΕΡΑ" And _
           UCase$(Trim$(CStr(ws.Cells(row, 3).Value))) = "ΩΡΑ" And _
           UCase$(Trim$(CStr(ws.Cells(row, 4).Value))) = "ΚΑΤΑΣΤΑΣΗ" Then
            FindPatientHeaderRow = row
            Exit Function
        End If
    Next row
End Function

Private Function WritePatchPayload(ByVal patchLines As Collection) As String
    Dim filePath As String, handle As Integer, i As Long
    filePath = Environ$("TEMP") & "\eka_replacement_patches_" & _
               Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    handle = FreeFile
    Open filePath For Output As #handle
    For i = 1 To patchLines.Count
        Print #handle, CStr(patchLines(i))
    Next i
    Close #handle
    WritePatchPayload = filePath
End Function

Private Sub RestoreReplacementTransaction( _
    ByVal rollbackPath As String, _
    ByVal replacementLines As Collection, _
    ByVal suspensionRows As Collection, _
    ByVal patchLines As Collection, _
    ByVal factCreated As Boolean, _
    ByVal factStartRow As Long, _
    ByVal factRows As Long)

    Dim backup As Workbook, parts As Variant, item As Variant
    Dim ws As Worksheet, oldAlerts As Boolean

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set backup = Application.Workbooks.Open(Filename:=rollbackPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    If backup.Windows.Count > 0 Then backup.Windows(1).Visible = False

    For Each item In patchLines
        parts = Split(CStr(item), vbTab)
        backup.Worksheets(PayloadValue(parts(1))).Range(PayloadValue(parts(2))).Copy _
            Destination:=ThisWorkbook.Worksheets(PayloadValue(parts(1))).Range(PayloadValue(parts(2)))
    Next item
    For Each item In suspensionRows
        backup.Worksheets("DAILY_INPUT").Range("A" & CStr(item) & ":E" & CStr(item)).Copy _
            Destination:=ThisWorkbook.Worksheets("DAILY_INPUT").Range("A" & CStr(item) & ":E" & CStr(item))
    Next item

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    If Not ws Is Nothing Then
        If factCreated Then
            ws.Delete
        ElseIf factRows > 0 Then
            ws.Unprotect
            ws.Range(ws.Cells(factStartRow, 1), ws.Cells(factStartRow + factRows - 1, 10)).ClearContents
            ws.Visible = xlSheetVeryHidden
            ws.Protect
        End If
    End If
    On Error GoTo 0

    Application.CutCopyMode = False
    backup.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
End Sub

Private Function ActiveReplacementCount() As Long
    Dim ws As Worksheet, row As Long, lastRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = "ACTIVE" Then _
            ActiveReplacementCount = ActiveReplacementCount + 1
    Next row
End Function

Private Sub WriteFactHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value2 = "ReplacementId"
    ws.Cells(1, 2).Value2 = "TargetSessionId"
    ws.Cells(1, 3).Value2 = "PatientId"
    ws.Cells(1, 4).Value2 = "OriginalTherapistId"
    ws.Cells(1, 5).Value2 = "ReplacementTherapistId"
    ws.Cells(1, 6).Value2 = "ReplacementDate"
    ws.Cells(1, 7).Value2 = "ReplacementTime"
    ws.Cells(1, 8).Value2 = "Reason"
    ws.Cells(1, 9).Value2 = "ProviderKind"
    ws.Cells(1, 10).Value2 = "Status"
End Sub

Private Function FactHeadersAreValid(ByVal ws As Worksheet) As Boolean
    Dim headers As Variant, i As Long
    headers = Array("ReplacementId", "TargetSessionId", "PatientId", "OriginalTherapistId", _
                    "ReplacementTherapistId", "ReplacementDate", "ReplacementTime", _
                    "Reason", "ProviderKind", "Status")
    For i = 0 To UBound(headers)
        If Trim$(CStr(ws.Cells(1, i + 1).Value)) <> CStr(headers(i)) Then Exit Function
    Next i
    FactHeadersAreValid = True
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

Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String) As String
    Dim marker As String, startPos As Long, index As Long
    Dim ch As String, escaped As Boolean, value As String
    marker = Chr$(34) & key & Chr$(34) & ":"
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos + Len(marker), jsonText, Chr$(34))
    If startPos = 0 Then Exit Function
    For index = startPos + 1 To Len(jsonText)
        ch = Mid$(jsonText, index, 1)
        If escaped Then
            Select Case ch
                Case "n": value = value & vbLf
                Case "r": value = value & vbCr
                Case "t": value = value & vbTab
                Case Else: value = value & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = Chr$(34) Then
            Exit For
        Else
            value = value & ch
        End If
    Next index
    JsonStringValue = value
End Function

Private Function ResolveReplacementPreparationScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\prepare_global_replacements_in_place.py"
    If Dir$(candidate) <> "" Then ResolveReplacementPreparationScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\prepare_global_replacements_in_place.py"
    If Dir$(candidate) <> "" Then _
        ResolveReplacementPreparationScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
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

Private Sub WriteReplacementDiagnostic( _
    ByVal errorNumber As Long, _
    ByVal errorText As String, _
    ByVal phase As String, _
    ByVal target As String)
    Dim filePath As String, handle As Integer
    filePath = Environ$("TEMP") & "\eka_replacement_last_error.log"
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #handle, "phase=" & phase
    Print #handle, "target=" & target
    Print #handle, "error_number=" & CStr(errorNumber)
    Print #handle, "error=" & errorText
    Close #handle
End Sub
'''


def install_global_replacement_action(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(GLOBAL_REPLACEMENT_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = GLOBAL_REPLACEMENT_MODULE_NAME
    module.CodeModule.AddFromString(GLOBAL_REPLACEMENT_MODULE_CODE)
