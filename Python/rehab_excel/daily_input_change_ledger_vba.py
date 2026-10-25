from __future__ import annotations


MODULE_NAME = "modDailyInputChangeLedger"
LEDGER_SHEET = "DAILY_CHANGE_LEDGER"
SHAPE_PREFIX = "ekaLedger_"

MODULE_CODE = r'''Option Explicit

Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const DAILY_INPUT_SHEET As String = "DAILY_INPUT"
Private Const DAILY_VIEW_SHEET As String = "THERAPIST_DAILY"
Private Const META_PREFIX As String = "DAILY_INPUT:"
Private Const STATUS_ACTIVE As String = "ACTIVE"
Private Const STATUS_RESTORED As String = "RESTORED"
Private Const SHAPE_PREFIX As String = "ekaLedger_"

Public Sub SyncDailyInputChangesToLedger()
    Dim inputWs As Worksheet, ledger As Worksheet
    Dim therapistHeader As Long, patientHeader As Long
    Dim row As Long, lastRow As Long
    Dim dateKey As String, displayName As String, reason As String
    Dim changeType As String, allDay As Boolean, timeText As String

    On Error GoTo SyncError
    Set inputWs = ThisWorkbook.Worksheets(DAILY_INPUT_SHEET)
    Set ledger = EnsureDailyChangeLedger()
    dateKey = Format$(inputWs.Range("B2").Value, "yyyy-mm-dd")
    therapistHeader = FindHeaderRow(inputWs, "Θεραπευτής")
    patientHeader = FindHeaderRow(inputWs, "Ασθενής")
    If therapistHeader = 0 Or patientHeader = 0 Or patientHeader <= therapistHeader Then Exit Sub

    For row = therapistHeader + 1 To patientHeader - 1
        displayName = Trim$(CStr(inputWs.Cells(row, 1).Value))
        If Len(displayName) > 0 Then
            allDay = IsYes(inputWs.Cells(row, 2).Value)
            timeText = IIf(allDay, "Όλη ημέρα", Trim$(CStr(inputWs.Cells(row, 3).Text)))
            reason = Trim$(CStr(inputWs.Cells(row, 5).Value))
            If Len(Trim$(CStr(inputWs.Cells(row, 6).Value))) > 0 Then _
                reason = JoinReason(reason, Trim$(CStr(inputWs.Cells(row, 6).Value)))
            AppendDailyInputLedgerRow ledger, dateKey, row, "therapist_absence", _
                                      displayName, timeText, reason
        End If
    Next row

    lastRow = inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).Row
    For row = patientHeader + 1 To lastRow
        displayName = Trim$(CStr(inputWs.Cells(row, 1).Value))
        If Len(displayName) = 0 Then GoTo NextPatient
        If IsPresentStatus(inputWs.Cells(row, 4).Value) Then GoTo NextPatient
        allDay = IsYes(inputWs.Cells(row, 2).Value)
        timeText = IIf(allDay, "Όλη ημέρα", Trim$(CStr(inputWs.Cells(row, 3).Text)))
        reason = Trim$(CStr(inputWs.Cells(row, 4).Value))
        If Len(Trim$(CStr(inputWs.Cells(row, 5).Value))) > 0 Then _
            reason = JoinReason(reason, Trim$(CStr(inputWs.Cells(row, 5).Value)))
        changeType = "patient_absence"
        If IsSessionCancellationStatus(inputWs.Cells(row, 4).Value) Then _
            changeType = "session_cancellation"
        AppendDailyInputLedgerRow ledger, dateKey, row, changeType, _
                                  displayName, timeText, reason
NextPatient:
    Next row

    ledger.Visible = xlSheetVeryHidden
    ThisWorkbook.Save
    On Error Resume Next
    Application.Run "RefreshDailyChangeLedgerView"
    RewireDailyInputRestoreButtons ledger
    EnsureTherapistNavigation
    On Error GoTo 0
    Exit Sub

SyncError:
    WriteRestoreDiagnostic Err.Number, Err.Description, "sync ledger", 0
End Sub

Public Sub RestoreDailyInputChangeRow()
    Dim callerName As String, ledgerRow As Long, sourceRow As Long
    Dim ledger As Worksheet, meta As String, sourceRowText As String
    Dim scriptPath As String, payloadPath As String, responsePath As String
    Dim commandLine As String, responseText As String, rollbackPath As String
    Dim exitCode As Long, previousEvents As Boolean, previousScreen As Boolean

    On Error GoTo RestoreError
    Set ledger = ThisWorkbook.Worksheets(LEDGER_SHEET)
    callerName = CStr(Application.Caller)
    ledgerRow = CLng(Mid$(callerName, InStrRev(callerName, "_") + 1))
    If UCase$(Trim$(CStr(ledger.Cells(ledgerRow, 12).Value))) <> STATUS_ACTIVE Then
        MsgBox "Η αλλαγή έχει ήδη επαναφερθεί.", vbInformation, "Επαναφορά"
        Exit Sub
    End If

    meta = Trim$(CStr(ledger.Cells(ledgerRow, 18).Value))
    If Left$(meta, Len(META_PREFIX)) <> META_PREFIX Then
        MsgBox "Η επιλεγμένη αλλαγή δεν είναι ημερήσια καταχώριση.", vbExclamation, "Επαναφορά"
        Exit Sub
    End If
    sourceRowText = Mid$(meta, InStrRev(meta, ":") + 1)
    If Not IsNumeric(sourceRowText) Then Err.Raise vbObjectError + 2501, , "Invalid source row metadata"
    sourceRow = CLng(sourceRowText)

    scriptPath = ResolveDailyInputRestoreScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Η λειτουργία επαναφοράς δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", _
               vbCritical, "Επαναφορά"
        Exit Sub
    End If

    ThisWorkbook.Save
    payloadPath = Environ$("TEMP") & "\eka_daily_input_restore_payload_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\eka_daily_input_restore_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --source-row " & CStr(sourceRow) & _
                  " --response " & QuoteArg(responsePath)
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or _
       InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or _
       Dir$(payloadPath) = "" Then
        MsgBox "Η επαναφορά δεν ολοκληρώθηκε. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Επαναφορά"
        GoTo CleanUp
    End If

    rollbackPath = Environ$("TEMP") & "\eka_daily_input_restore_rollback_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".xlsm"
    ThisWorkbook.SaveCopyAs rollbackPath

    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ResetMasterDailyPresentation
    ThisWorkbook.Worksheets(DAILY_INPUT_SHEET).Range("A" & CStr(sourceRow) & ":F" & CStr(sourceRow)).ClearContents

    If Not ApplyValidatedDailyInputPayload(payloadPath) Then
        RestoreDailyInputTransaction rollbackPath, sourceRow, ledgerRow
        GoTo RestoreExit
    End If

    ledger.Cells(ledgerRow, 12).Value = STATUS_RESTORED
    ledger.Cells(ledgerRow, 17).Value = Now
    ThisWorkbook.Save
    RenderTherapistWorkloadStrip
    SyncDailyInputChangesToLedger
    MsgBox "Η αλλαγή επαναφέρθηκε και το ημερήσιο πρόγραμμα ενημερώθηκε.", _
           vbInformation, "Επαναφορά"

RestoreExit:
    Application.ScreenUpdating = previousScreen
    Application.EnableEvents = previousEvents
    GoTo CleanUp

RestoreError:
    WriteRestoreDiagnostic Err.Number, Err.Description, "restore daily input change", ledgerRow
    On Error Resume Next
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then _
        RestoreDailyInputTransaction rollbackPath, sourceRow, ledgerRow
    On Error GoTo 0
    MsgBox "Η επαναφορά δεν ολοκληρώθηκε. Επικοινωνήστε με τον διαχειριστή.", _
           vbCritical, "Επαναφορά"
    Resume RestoreExit

CleanUp:
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then Kill rollbackPath
    On Error GoTo 0
End Sub

Private Function EnsureDailyChangeLedger() As Worksheet
    Dim ws As Worksheet, daily As Worksheet, startRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LEDGER_SHEET
        WriteLedgerHeaders ws
        Set daily = ThisWorkbook.Worksheets(DAILY_VIEW_SHEET)
        startRow = Application.Max(21, daily.UsedRange.Row + daily.UsedRange.Rows.Count + 2)
        ws.Range("R1").Value = startRow
    ElseIf Len(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 Then
        WriteLedgerHeaders ws
    End If
    ws.Visible = xlSheetVeryHidden
    Set EnsureDailyChangeLedger = ws
End Function

Private Sub WriteLedgerHeaders(ByVal ws As Worksheet)
    Dim headers As Variant, i As Long
    headers = Array("ChangeID", "Date", "PatientID", "Patient", "OriginalProvider", "OriginalTime", _
                    "EffectiveProvider", "EffectiveTime", "Reason", "OriginalCell", "EffectiveCell", _
                    "Status", "OriginalSnapshot", "EffectiveSnapshot", "ExpectedOriginal", "ExpectedEffective", _
                    "RestoredAt", "Reserved", "DisplayRow")
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
End Sub

Private Sub AppendDailyInputLedgerRow( _
    ByVal ledger As Worksheet, ByVal dateKey As String, ByVal sourceRow As Long, _
    ByVal changeType As String, ByVal displayName As String, ByVal timeText As String, _
    ByVal reason As String)

    Dim meta As String, row As Long, lastRow As Long, targetRow As Long
    Dim baseId As String, changeId As String, suffix As Long
    meta = META_PREFIX & changeType & ":" & CStr(sourceRow)
    lastRow = ledger.Cells(ledger.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    For row = 2 To lastRow
        If Trim$(CStr(ledger.Cells(row, 18).Value)) = meta And _
           UCase$(Trim$(CStr(ledger.Cells(row, 12).Value))) = STATUS_ACTIVE Then Exit Sub
    Next row

    baseId = "daily-input:" & dateKey & ":" & changeType & ":" & CStr(sourceRow)
    changeId = baseId
    suffix = 1
    Do While LedgerIdExists(ledger, changeId)
        suffix = suffix + 1
        changeId = baseId & ":v" & CStr(suffix)
    Loop

    targetRow = lastRow + 1
    ledger.Cells(targetRow, 1).Value = changeId
    ledger.Cells(targetRow, 2).Value = dateKey
    ledger.Cells(targetRow, 4).Value = displayName
    ledger.Cells(targetRow, 5).Value = "Ημερήσια κατάσταση"
    ledger.Cells(targetRow, 6).Value = timeText
    ledger.Cells(targetRow, 7).Value = "Εκτός προγράμματος"
    ledger.Cells(targetRow, 9).Value = reason
    ledger.Cells(targetRow, 12).Value = STATUS_ACTIVE
    ledger.Cells(targetRow, 18).Value = meta
End Sub

Private Sub RewireDailyInputRestoreButtons(ByVal ledger As Worksheet)
    Dim row As Long, lastRow As Long, meta As String, shapeName As String
    Dim wsDaily As Worksheet, shp As Shape
    Set wsDaily = ThisWorkbook.Worksheets(DAILY_VIEW_SHEET)
    lastRow = ledger.Cells(ledger.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        meta = Trim$(CStr(ledger.Cells(row, 18).Value))
        If Left$(meta, Len(META_PREFIX)) = META_PREFIX And _
           UCase$(Trim$(CStr(ledger.Cells(row, 12).Value))) = STATUS_ACTIVE Then
            shapeName = SHAPE_PREFIX & "Restore_" & CStr(row)
            On Error Resume Next
            Set shp = wsDaily.Shapes(shapeName)
            If Not shp Is Nothing Then shp.OnAction = "RestoreDailyInputChangeRow"
            wsDaily.Shapes(SHAPE_PREFIX & "Select_" & CStr(row)).Delete
            Set shp = Nothing
            On Error GoTo 0
        End If
    Next row
End Sub

Private Sub ResetMasterDailyPresentation()
    Dim master As Worksheet
    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    With master.Range("D2:J5000").Font
        .Strikethrough = False
        .Color = RGB(0, 0, 0)
    End With
    With master.Range("M2:M5000").Font
        .Strikethrough = False
        .Color = RGB(0, 0, 0)
    End With
End Sub

Private Sub RestoreDailyInputTransaction(ByVal rollbackPath As String, ByVal sourceRow As Long, ByVal ledgerRow As Long)
    Dim backup As Workbook, oldAlerts As Boolean
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set backup = Application.Workbooks.Open(rollbackPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error Resume Next
    If backup.Windows.Count > 0 Then backup.Windows(1).Visible = False
    backup.Worksheets("MASTER_SCHEDULE").Range("A1:M5000").Copy _
        Destination:=ThisWorkbook.Worksheets("MASTER_SCHEDULE").Range("A1:M5000")
    backup.Worksheets("THERAPIST_DAILY").Range("A1:J200").Copy _
        Destination:=ThisWorkbook.Worksheets("THERAPIST_DAILY").Range("A1:J200")
    backup.Worksheets(DAILY_INPUT_SHEET).Range("A" & CStr(sourceRow) & ":F" & CStr(sourceRow)).Copy _
        Destination:=ThisWorkbook.Worksheets(DAILY_INPUT_SHEET).Range("A" & CStr(sourceRow) & ":F" & CStr(sourceRow))
    backup.Worksheets(LEDGER_SHEET).Range("A" & CStr(ledgerRow) & ":S" & CStr(ledgerRow)).Copy _
        Destination:=ThisWorkbook.Worksheets(LEDGER_SHEET).Range("A" & CStr(ledgerRow) & ":S" & CStr(ledgerRow))
    Application.CutCopyMode = False
    backup.Close SaveChanges:=False
    ThisWorkbook.Save
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim row As Long
    For row = 1 To 100
        If Trim$(CStr(ws.Cells(row, 1).Value)) = headerText Then
            FindHeaderRow = row
            Exit Function
        End If
    Next row
End Function

Private Function LedgerIdExists(ByVal ledger As Worksheet, ByVal changeId As String) As Boolean
    Dim found As Range
    Set found = ledger.Columns(1).Find(What:=changeId, LookAt:=xlWhole, MatchCase:=False)
    LedgerIdExists = Not found Is Nothing
End Function

Private Function IsYes(ByVal value As Variant) As Boolean
    Dim text As String
    text = UCase$(Trim$(CStr(value)))
    IsYes = (text = "ΝΑΙ" Or text = "NAI" Or text = "YES" Or text = "Y" Or text = "1" Or text = "TRUE")
End Function

Private Function IsPresentStatus(ByVal value As Variant) As Boolean
    Dim text As String
    text = UCase$(Trim$(CStr(value)))
    IsPresentStatus = (text = "ΠΑΡΩΝ" Or text = "PARON" Or text = "PRESENT")
End Function

Private Function IsSessionCancellationStatus(ByVal value As Variant) As Boolean
    Dim text As String
    text = UCase$(Trim$(CStr(value)))
    IsSessionCancellationStatus = _
        (text = "ΑΝΑΒΟΛΗ ΤΜΗΜΑΤΟΣ" Or text = "ΔΕΝ ΠΡΟΣΗΛΘΕ" Or text = "ΑΝΑΣΤΟΛΗ ΘΕΡΑΠΕΙΑΣ")
End Function

Private Function JoinReason(ByVal firstText As String, ByVal secondText As String) As String
    If Len(firstText) = 0 Then
        JoinReason = secondText
    ElseIf Len(secondText) = 0 Then
        JoinReason = firstText
    Else
        JoinReason = firstText & " | " & secondText
    End If
End Function

Private Function ResolveDailyInputRestoreScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\prepare_daily_input_restore_in_place.py"
    If Dir$(candidate) <> "" Then ResolveDailyInputRestoreScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\prepare_daily_input_restore_in_place.py"
    If Dir$(candidate) <> "" Then _
        ResolveDailyInputRestoreScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
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

Private Sub WriteRestoreDiagnostic(ByVal errorNumber As Long, ByVal errorText As String, ByVal phase As String, ByVal ledgerRow As Long)
    Dim filePath As String, handle As Integer
    filePath = Environ$("TEMP") & "\eka_daily_input_restore_last_error.log"
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #handle, "phase=" & phase
    Print #handle, "ledger_row=" & CStr(ledgerRow)
    Print #handle, "error_number=" & CStr(errorNumber)
    Print #handle, "error=" & errorText
    Close #handle
End Sub
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def install_daily_input_change_ledger(vbproject) -> None:
    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)
