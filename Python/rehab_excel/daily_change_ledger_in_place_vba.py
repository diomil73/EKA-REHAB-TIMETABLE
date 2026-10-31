from __future__ import annotations


MODULE_NAME = "modDailyChangeLedger"
LEDGER_SHEET = "DAILY_CHANGE_LEDGER"
HEADERS = (
    "ChangeID",
    "Date",
    "ChangeType",
    "SubjectID",
    "Subject",
    "FromText",
    "ToText",
    "Reason",
    "SourceRow",
    "Status",
    "RestoredAt",
    "DisplayRow",
)


VBA_CODE = r'''Option Explicit

Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const DAILY_SHEET As String = "THERAPIST_DAILY"
Private Const INPUT_SHEET As String = "DAILY_INPUT"
Private Const SHAPE_PREFIX As String = "ekaLedger_"
Private Const STATUS_ACTIVE As String = "ACTIVE"
Private Const STATUS_RESTORED As String = "RESTORED"

Public Sub ValidateDailyChangePayload(ByVal parts As Variant)
    If UBound(parts) <> 9 Then Err.Raise vbObjectError + 2801, , "Invalid CHANGE payload"
    If PayloadValue(parts(0)) <> "CHANGE" Then Err.Raise vbObjectError + 2802, , "Invalid CHANGE row kind"
    If Len(PayloadValue(parts(1))) = 0 Then Err.Raise vbObjectError + 2803, , "Missing change id"
    If Not IsNumeric(PayloadValue(parts(9))) Then Err.Raise vbObjectError + 2804, , "Invalid source row"
End Sub

Public Sub PersistDailyChangePayload(ByVal parts As Variant)
    Dim ws As Worksheet, lastRow As Long, row As Long, changeId As String
    ValidateDailyChangePayload parts
    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    changeId = PayloadValue(parts(1))
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = changeId Then
            If UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = STATUS_ACTIVE Then Exit Sub
            Err.Raise vbObjectError + 2805, , "Daily change id already exists with non-active state"
        End If
    Next row
    row = IIf(lastRow < 2, 2, lastRow + 1)
    ws.Cells(row, 1).Value2 = changeId
    ws.Cells(row, 2).Value2 = PayloadValue(parts(2))
    ws.Cells(row, 3).Value2 = PayloadValue(parts(3))
    ws.Cells(row, 4).Value2 = PayloadValue(parts(4))
    ws.Cells(row, 5).Value2 = PayloadValue(parts(5))
    ws.Cells(row, 6).Value2 = PayloadValue(parts(6))
    ws.Cells(row, 7).Value2 = PayloadValue(parts(7))
    ws.Cells(row, 8).Value2 = PayloadValue(parts(8))
    ws.Cells(row, 9).Value2 = CLng(PayloadValue(parts(9)))
    ws.Cells(row, 10).Value2 = STATUS_ACTIVE
    ws.Visible = xlSheetVeryHidden
End Sub

Public Sub RefreshDailyChangeLedgerView()
    Dim wsDaily As Worksheet, wsLedger As Worksheet
    Dim startRow As Long, lastLedgerRow As Long, ledgerRow As Long, displayRow As Long
    Dim shp As Shape, cb As Shape, btn As Shape
    Dim statusText As String, changeType As String

    On Error GoTo RenderError
    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)

    For Each shp In wsDaily.Shapes
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then shp.Delete
    Next shp

    startRow = wsDaily.Cells(wsDaily.Rows.Count, 1).End(xlUp).Row + 3
    If startRow < 21 Then startRow = 21
    wsDaily.Range("A" & startRow & ":G" & startRow + 200).ClearContents
    wsDaily.Range("A" & startRow & ":G" & startRow + 200).Font.Italic = False
    wsDaily.Range("A" & startRow & ":G" & startRow + 200).Font.Bold = False

    lastLedgerRow = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row
    If lastLedgerRow < 2 Then Exit Sub

    wsDaily.Cells(startRow, 1).Value = "Αλλαγές ημέρας"
    With wsDaily.Range("A" & startRow & ":G" & startRow)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    wsDaily.Cells(startRow + 1, 1).Value = "✓"
    wsDaily.Cells(startRow + 1, 2).Value = "Ασθενής / Θεραπευτής"
    wsDaily.Cells(startRow + 1, 3).Value = "Από"
    wsDaily.Cells(startRow + 1, 4).Value = "Προς"
    wsDaily.Cells(startRow + 1, 5).Value = "Αιτία"
    wsDaily.Cells(startRow + 1, 6).Value = "Κατάσταση"
    wsDaily.Cells(startRow + 1, 7).Value = "Ενέργεια"
    wsDaily.Range("A" & startRow + 1 & ":G" & startRow + 1).Font.Bold = True

    displayRow = startRow + 2
    For ledgerRow = 2 To lastLedgerRow
        If Len(Trim$(CStr(wsLedger.Cells(ledgerRow, 1).Value))) = 0 Then GoTo NextLedger
        statusText = UCase$(Trim$(CStr(wsLedger.Cells(ledgerRow, 10).Value)))
        changeType = LCase$(Trim$(CStr(wsLedger.Cells(ledgerRow, 3).Value)))
        wsDaily.Cells(displayRow, 2).Value = wsLedger.Cells(ledgerRow, 5).Value
        wsDaily.Cells(displayRow, 3).Value = wsLedger.Cells(ledgerRow, 6).Value
        wsDaily.Cells(displayRow, 4).Value = wsLedger.Cells(ledgerRow, 7).Value
        wsDaily.Cells(displayRow, 5).Value = GreekLedgerReason(CStr(wsLedger.Cells(ledgerRow, 8).Value))
        If statusText = STATUS_RESTORED Then
            wsDaily.Cells(displayRow, 6).Value = "Επαναφέρθηκε"
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.Color = RGB(128, 128, 128)
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.Italic = True
        Else
            wsDaily.Cells(displayRow, 6).Value = "Ενεργή αλλαγή"
            If changeType = "patient_absence" Or changeType = "session_cancellation" Then
                Set cb = wsDaily.Shapes.AddFormControl(1, wsDaily.Cells(displayRow, 1).Left + 4, wsDaily.Cells(displayRow, 1).Top + 2, 16, 16)
                cb.Name = SHAPE_PREFIX & "Select_" & CStr(ledgerRow)
                cb.TextFrame.Characters.Text = ""
                Set btn = wsDaily.Shapes.AddFormControl(0, wsDaily.Cells(displayRow, 7).Left + 2, wsDaily.Cells(displayRow, 7).Top + 1, wsDaily.Cells(displayRow, 7).Width - 4, wsDaily.Cells(displayRow, 7).Height - 2)
                btn.Name = SHAPE_PREFIX & "Restore_" & CStr(ledgerRow)
                btn.TextFrame.Characters.Text = "Επαναφορά"
                btn.OnAction = "RestoreDailyChangeRow"
            Else
                wsDaily.Cells(displayRow, 7).Value = "Επαναφορά σύντομα"
            End If
        End If
        wsLedger.Cells(ledgerRow, 12).Value2 = displayRow
        displayRow = displayRow + 1
NextLedger:
    Next ledgerRow

    Set btn = wsDaily.Shapes.AddFormControl(0, wsDaily.Cells(displayRow + 1, 2).Left, wsDaily.Cells(displayRow + 1, 2).Top, 170, 26)
    btn.Name = SHAPE_PREFIX & "RestoreSelected"
    btn.TextFrame.Characters.Text = "Επαναφορά επιλεγμένων"
    btn.OnAction = "RestoreSelectedDailyChanges"

    Set btn = wsDaily.Shapes.AddFormControl(0, wsDaily.Cells(displayRow + 1, 5).Left, wsDaily.Cells(displayRow + 1, 5).Top, 145, 26)
    btn.Name = SHAPE_PREFIX & "RestoreAll"
    btn.TextFrame.Characters.Text = "Επαναφορά όλων"
    btn.OnAction = "RestoreAllDailyChanges"

    wsDaily.Columns("A").ColumnWidth = 8
    wsDaily.Columns("B").ColumnWidth = 28
    wsDaily.Columns("C:D").ColumnWidth = 24
    wsDaily.Columns("E").ColumnWidth = 30
    wsDaily.Columns("F").ColumnWidth = 18
    wsDaily.Columns("G").ColumnWidth = 21
    Exit Sub
RenderError:
    MsgBox "Δεν ήταν δυνατή η εμφάνιση της λίστας αλλαγών: " & Err.Description, vbExclamation, "Αλλαγές ημέρας"
End Sub

Public Sub RestoreDailyChangeRow()
    Dim callerName As String, ledgerRow As Long, changeId As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    callerName = CStr(Application.Caller)
    ledgerRow = CLng(Mid$(callerName, InStrRev(callerName, "_") + 1))
    If UCase$(Trim$(CStr(ws.Cells(ledgerRow, 10).Value))) <> STATUS_ACTIVE Then Exit Sub
    changeId = Trim$(CStr(ws.Cells(ledgerRow, 1).Value))
    If Len(changeId) = 0 Then Exit Sub
    RunDailyChangeRestore " --change-id " & QuoteArg(changeId)
End Sub

Public Sub RestoreSelectedDailyChanges()
    Dim wsDaily As Worksheet, wsLedger As Worksheet, shp As Shape
    Dim ledgerRow As Long, changeId As String, args As String, selectedCount As Long
    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)
    For Each shp In wsDaily.Shapes
        If Left$(shp.Name, Len(SHAPE_PREFIX & "Select_")) = SHAPE_PREFIX & "Select_" Then
            If shp.ControlFormat.Value = 1 Then
                ledgerRow = CLng(Mid$(shp.Name, Len(SHAPE_PREFIX & "Select_") + 1))
                If UCase$(Trim$(CStr(wsLedger.Cells(ledgerRow, 10).Value))) = STATUS_ACTIVE Then
                    changeId = Trim$(CStr(wsLedger.Cells(ledgerRow, 1).Value))
                    If Len(changeId) > 0 Then args = args & " --change-id " & QuoteArg(changeId): selectedCount = selectedCount + 1
                End If
            End If
        End If
    Next shp
    If selectedCount = 0 Then
        MsgBox "Δεν επιλέχθηκε ενεργή αλλαγή για επαναφορά.", vbInformation, "Επαναφορά επιλεγμένων"
        Exit Sub
    End If
    RunDailyChangeRestore args
End Sub

Public Sub RestoreAllDailyChanges()
    Dim ws As Worksheet, lastRow As Long, row As Long, changeType As String, args As String, count As Long
    If MsgBox("Να επαναφερθούν όλες οι διαθέσιμες ενεργές αλλαγές της ημέρας;", vbYesNo + vbQuestion, "Επαναφορά όλων") <> vbYes Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        changeType = LCase$(Trim$(CStr(ws.Cells(row, 3).Value)))
        If UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = STATUS_ACTIVE And _
           (changeType = "patient_absence" Or changeType = "session_cancellation") Then
            args = args & " --change-id " & QuoteArg(Trim$(CStr(ws.Cells(row, 1).Value)))
            count = count + 1
        End If
    Next row
    If count = 0 Then
        MsgBox "Δεν υπάρχουν διαθέσιμες ενεργές αλλαγές για επαναφορά.", vbInformation, "Επαναφορά όλων"
        Exit Sub
    End If
    RunDailyChangeRestore args
End Sub

Private Function RunDailyChangeRestore(ByVal changeArgs As String) As Boolean
    Dim scriptPath As String, payloadPath As String, responsePath As String, commandLine As String
    Dim responseText As String, exitCode As Long
    On Error GoTo Failed
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 2820, , "Το αρχείο είναι μόνο για ανάγνωση."
    scriptPath = ResolveDailyChangeRestoreScript()
    If Len(scriptPath) = 0 Then Err.Raise vbObjectError + 2821, , "Δεν βρέθηκε το εργαλείο επαναφοράς."
    ThisWorkbook.Save
    payloadPath = Environ$("TEMP") & "\eka_daily_restore_payload_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\eka_daily_restore_response_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & " --input " & QuoteArg(ThisWorkbook.FullName) & " --payload " & QuoteArg(payloadPath) & " --response " & QuoteArg(responsePath) & changeArgs
    Application.StatusBar = "Προετοιμασία επαναφοράς..."
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    Application.StatusBar = False
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then
        Err.Raise vbObjectError + 2822, , "Η προετοιμασία επαναφοράς απέτυχε."
    End If
    If Not ApplyDailyChangeRestorePayload(payloadPath) Then GoTo CleanUp
    RefreshDailyChangeLedgerView
    MsgBox "Η επαναφορά ολοκληρώθηκε επιτυχώς.", vbInformation, "Επαναφορά"
    RunDailyChangeRestore = True
CleanUp:
    Application.StatusBar = False
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
    Exit Function
Failed:
    Application.StatusBar = False
    WriteRestoreDiagnostic Err.Number, Err.Description, "restore action"
    MsgBox "Η επαναφορά δεν ολοκληρώθηκε: " & Err.Description, vbExclamation, "Επαναφορά"
    Resume CleanUp
End Function

Public Function ApplyDailyChangeRestorePayload(ByVal payloadPath As String) As Boolean
    Dim rawText As String, lines As Variant, parts As Variant, i As Long
    Dim rollbackPath As String, patchPath As String, patchText As String
    Dim ws As Worksheet, rowNumber As Long, changeId As String
    Dim previousEvents As Boolean, previousScreenUpdating As Boolean
    On Error GoTo Failed
    rawText = ReadUtf8Text(payloadPath)
    lines = Split(Replace(rawText, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo PreflightNext
        parts = Split(CStr(lines(i)), vbTab)
        Select Case CStr(parts(0))
            Case "SOURCE_CLEAR"
                If UBound(parts) <> 4 Then Err.Raise vbObjectError + 2830, , "Invalid SOURCE_CLEAR payload"
                If PayloadValue(parts(1)) <> INPUT_SHEET Or PayloadValue(parts(3)) <> "A" Or PayloadValue(parts(4)) <> "E" Then Err.Raise vbObjectError + 2831, , "Unsafe source clear target"
                If Not IsNumeric(PayloadValue(parts(2))) Or CLng(PayloadValue(parts(2))) < 2 Then Err.Raise vbObjectError + 2832, , "Invalid source row"
            Case "RESTORE_CHANGE"
                If UBound(parts) <> 1 Then Err.Raise vbObjectError + 2833, , "Invalid RESTORE_CHANGE payload"
                If Not ActiveLedgerChangeExists(PayloadValue(parts(1))) Then Err.Raise vbObjectError + 2834, , "Restore change is not ACTIVE"
            Case "PATCH"
                If UBound(parts) <> 12 Then Err.Raise vbObjectError + 2835, , "Invalid PATCH payload"
                patchText = patchText & CStr(lines(i)) & vbLf
            Case Else
                Err.Raise vbObjectError + 2836, , "Unsupported restore payload row"
        End Select
PreflightNext:
    Next i
    If Len(patchText) = 0 Then Err.Raise vbObjectError + 2837, , "Restore payload has no projection patches"

    rollbackPath = Environ$("TEMP") & "\eka_daily_restore_rollback_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".xlsm"
    ThisWorkbook.SaveCopyAs rollbackPath
    patchPath = Environ$("TEMP") & "\eka_daily_restore_patches_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    WriteUtf8Text patchPath, patchText

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo ApplyNext
        parts = Split(CStr(lines(i)), vbTab)
        If CStr(parts(0)) = "SOURCE_CLEAR" Then
            rowNumber = CLng(PayloadValue(parts(2)))
            ws.Range("A" & rowNumber & ":E" & rowNumber).ClearContents
        End If
ApplyNext:
    Next i
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating

    If Not ApplyValidatedDailyInputPayload(patchPath) Then GoTo Failed

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo MarkNext
        parts = Split(CStr(lines(i)), vbTab)
        If CStr(parts(0)) = "RESTORE_CHANGE" Then
            changeId = PayloadValue(parts(1))
            MarkLedgerChangeRestored changeId
        End If
MarkNext:
    Next i
    ThisWorkbook.Save
    ApplyDailyChangeRestorePayload = True
    GoTo CleanUp

Failed:
    WriteRestoreDiagnostic Err.Number, Err.Description, "apply restore payload"
    On Error Resume Next
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then RestoreDailyChangeTransaction rollbackPath
    ThisWorkbook.Save
    On Error GoTo 0
    MsgBox "Η επαναφορά απέτυχε και οι αλλαγές αναιρέθηκαν.", vbCritical, "Επαναφορά"
CleanUp:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    On Error Resume Next
    If Len(patchPath) > 0 And Dir$(patchPath) <> "" Then Kill patchPath
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then Kill rollbackPath
    On Error GoTo 0
End Function

Private Function ActiveLedgerChangeExists(ByVal changeId As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, row As Long
    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = changeId And UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = STATUS_ACTIVE Then ActiveLedgerChangeExists = True: Exit Function
    Next row
End Function

Private Sub MarkLedgerChangeRestored(ByVal changeId As String)
    Dim ws As Worksheet, lastRow As Long, row As Long
    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = changeId And UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = STATUS_ACTIVE Then
            ws.Cells(row, 10).Value2 = STATUS_RESTORED
            ws.Cells(row, 11).Value = Now
            Exit Sub
        End If
    Next row
    Err.Raise vbObjectError + 2840, , "Ledger change not found during restore"
End Sub

Public Sub RestoreDailyChangeLedgerFromRollback(ByVal rollbackPath As String)
    RestoreOneSheetFromRollback rollbackPath, LEDGER_SHEET
End Sub

Private Sub RestoreDailyChangeTransaction(ByVal rollbackPath As String)
    RestoreOneSheetFromRollback rollbackPath, INPUT_SHEET
    RestoreOneSheetFromRollback rollbackPath, LEDGER_SHEET
    RestoreOneSheetFromRollback rollbackPath, "THERAPIST_DAILY"
    RestoreOneSheetFromRollback rollbackPath, "MASTER_SCHEDULE"
End Sub

Private Sub RestoreOneSheetFromRollback(ByVal rollbackPath As String, ByVal sheetName As String)
    Dim rollbackBook As Workbook, sourceSheet As Worksheet, targetSheet As Worksheet
    Dim previousAlerts As Boolean
    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set rollbackBook = Application.Workbooks.Open(Filename:=rollbackPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set sourceSheet = rollbackBook.Worksheets(sheetName)
    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    targetSheet.Cells.Clear
    sourceSheet.UsedRange.Copy Destination:=targetSheet.Range(sourceSheet.UsedRange.Address)
    Application.CutCopyMode = False
    rollbackBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
End Sub

Private Function GreekLedgerReason(ByVal reasonText As String) As String
    Select Case LCase$(Trim$(reasonText))
        Case "", "temporary replacement", "daily_input therapist absence", "daily input therapist absence"
            GreekLedgerReason = "Απουσία θεραπευτή"
        Case Else
            GreekLedgerReason = reasonText
    End Select
End Function

Private Function ResolveDailyChangeRestoreScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\prepare_daily_change_restore_in_place.py"
    If Dir$(candidate) <> "" Then ResolveDailyChangeRestoreScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\prepare_daily_change_restore_in_place.py"
    If Dir$(candidate) <> "" Then ResolveDailyChangeRestoreScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
End Function

Private Sub WriteRestoreDiagnostic(ByVal errorNumber As Long, ByVal errorText As String, ByVal phase As String)
    Dim filePath As String, handle As Integer
    filePath = Environ$("TEMP") & "\eka_daily_restore_last_error.log"
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #handle, "phase=" & phase
    Print #handle, "error_number=" & CStr(errorNumber)
    Print #handle, "error=" & errorText
    Close #handle
End Sub

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

Private Sub WriteUtf8Text(ByVal filePath As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _replace_all(component, code: str) -> None:
    module = component.CodeModule
    count = int(module.CountOfLines)
    if count:
        module.DeleteLines(1, count)
    module.AddFromString(code)


def patch_daily_input_for_ledger(code: str) -> str:
    normalized = code.replace("\r\n", "\n").replace("\r", "\n")
    old_preflight = '''        parts = Split(CStr(lines(i)), vbTab)
        If UBound(parts) <> 12 Or CStr(parts(0)) <> "PATCH" Then Err.Raise vbObjectError + 2301, , "Invalid DAILY_INPUT payload"
        currentPatchIndex = patchCount + 1
        currentSheet = PayloadValue(parts(1))
        currentCell = PayloadValue(parts(2))
        ValidatePatchTarget parts
        patchCount = patchCount + 1'''
    new_preflight = '''        parts = Split(CStr(lines(i)), vbTab)
        If CStr(parts(0)) = "CHANGE" Then
            ValidateDailyChangePayload parts
        ElseIf CStr(parts(0)) = "PATCH" Then
            If UBound(parts) <> 12 Then Err.Raise vbObjectError + 2301, , "Invalid DAILY_INPUT patch payload"
            currentPatchIndex = patchCount + 1
            currentSheet = PayloadValue(parts(1))
            currentCell = PayloadValue(parts(2))
            ValidatePatchTarget parts
            patchCount = patchCount + 1
        Else
            Err.Raise vbObjectError + 2301, , "Invalid DAILY_INPUT payload"
        End If'''
    if old_preflight not in normalized:
        raise RuntimeError("DAILY_INPUT ledger preflight marker not found")
    normalized = normalized.replace(old_preflight, new_preflight, 1)

    old_apply = '''        parts = Split(CStr(lines(i)), vbTab)
        currentPatchIndex = currentPatchIndex + 1
        currentSheet = PayloadValue(parts(1))
        currentCell = PayloadValue(parts(2))
        currentPhase = "apply patch"
        ApplyPatch parts
        currentPhase = "verify patch"
        If Not VerifyPatch(parts) Then Err.Raise vbObjectError + 2303, , "DAILY_INPUT read-back mismatch"'''
    new_apply = '''        parts = Split(CStr(lines(i)), vbTab)
        If CStr(parts(0)) = "CHANGE" Then
            currentPhase = "record daily change"
            PersistDailyChangePayload parts
        Else
            currentPatchIndex = currentPatchIndex + 1
            currentSheet = PayloadValue(parts(1))
            currentCell = PayloadValue(parts(2))
            currentPhase = "apply patch"
            ApplyPatch parts
            currentPhase = "verify patch"
            If Not VerifyPatch(parts) Then Err.Raise vbObjectError + 2303, , "DAILY_INPUT read-back mismatch"
        End If'''
    if old_apply not in normalized:
        raise RuntimeError("DAILY_INPUT ledger apply marker not found")
    normalized = normalized.replace(old_apply, new_apply, 1)

    rollback_marker = '''        parts = Split(CStr(lines(i)), vbTab)
        sheetName = PayloadValue(parts(1))'''
    rollback_replacement = '''        parts = Split(CStr(lines(i)), vbTab)
        If CStr(parts(0)) <> "PATCH" Then GoTo RestoreNext
        sheetName = PayloadValue(parts(1))'''
    if rollback_marker not in normalized:
        raise RuntimeError("DAILY_INPUT rollback marker not found")
    normalized = normalized.replace(rollback_marker, rollback_replacement, 1)

    error_marker = '''    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then
        RestorePatchesFromRollback rollbackPath, lines
        ThisWorkbook.Save
    End If'''
    error_replacement = '''    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then
        RestorePatchesFromRollback rollbackPath, lines
        RestoreDailyChangeLedgerFromRollback rollbackPath
        ThisWorkbook.Save
    End If'''
    if error_marker not in normalized:
        raise RuntimeError("DAILY_INPUT ledger rollback integration marker not found")
    normalized = normalized.replace(error_marker, error_replacement, 1)

    refresh_marker = '''    ApplyMasterView
    RefreshOperationalViews = True'''
    refresh_replacement = '''    ApplyMasterView
    RefreshDailyChangeLedgerView
    RefreshOperationalViews = True'''
    if refresh_marker not in normalized:
        raise RuntimeError("DAILY_INPUT ledger refresh marker not found")
    return normalized.replace(refresh_marker, refresh_replacement, 1)


def ensure_daily_change_ledger_sheet(workbook) -> None:
    try:
        ws = workbook.Worksheets(LEDGER_SHEET)
        ws.Cells.Clear()
    except Exception:
        ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        ws.Name = LEDGER_SHEET
    for col, header in enumerate(HEADERS, start=1):
        ws.Cells(1, col).Value2 = header
    ws.Visible = 2


def install_daily_change_ledger(vbproject, workbook) -> None:
    ensure_daily_change_ledger_sheet(workbook)
    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(VBA_CODE)

    daily = vbproject.VBComponents("modDailyInputAction")
    current = str(daily.CodeModule.Lines(1, int(daily.CodeModule.CountOfLines)))
    _replace_all(daily, patch_daily_input_for_ledger(current))
