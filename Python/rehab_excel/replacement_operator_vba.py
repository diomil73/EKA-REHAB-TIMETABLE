from __future__ import annotations


MODULE_NAME = "modReplacementOperator"
FORM_NAME = "frmReplacementDecision"
MACRO_NAME = "RunReplacementOperatorWorkflow"
DAILY_INPUT_MODULE_NAME = "modDailyInputAction"


MODULE_CODE = r'''Option Explicit

Private Const FACT_SHEET As String = "DAILY_REPLACEMENT_FACTS"
Private mRowsPath As String
Private mDecisionResponse As String
Private mDecisionAccepted As Boolean
Private mWorkflowCancelled As Boolean

Public Function RunReplacementOperatorWorkflow() As Boolean
    Dim iteration As Long
    Dim mode As String

    On Error GoTo WorkflowError

    For iteration = 1 To 100
        If Not PrepareNextReplacementDecision() Then GoTo WorkflowExit
        mode = ReplacementJsonStringValue(mDecisionResponse, "mode")

        If mode = "complete" Then
            RunReplacementOperatorWorkflow = True
            GoTo WorkflowExit
        End If

        If mode <> "decision_required" Then
            MsgBox "Η διαδικασία αναπλήρωσης επέστρεψε μη αναμενόμενη κατάσταση.", _
                   vbCritical, "Αναπληρώσεις"
            GoTo WorkflowExit
        End If

        mDecisionAccepted = False
        mWorkflowCancelled = False
        frmReplacementDecision.LoadDecision mRowsPath, mDecisionResponse
        frmReplacementDecision.Show vbModal

        If mWorkflowCancelled Then GoTo WorkflowExit
        If Not mDecisionAccepted Then GoTo WorkflowExit
    Next iteration

    MsgBox "Η διαδικασία αναπλήρωσης σταμάτησε για λόγους ασφαλείας μετά από πολλές διαδοχικές αποφάσεις.", _
           vbExclamation, "Αναπληρώσεις"
    GoTo WorkflowExit

WorkflowError:
    WriteOperatorDiagnostic Err.Number, Err.Description, "operator workflow"
    MsgBox "Η διαδικασία αναπλήρωσης δεν ολοκληρώθηκε. Οι ήδη επιβεβαιωμένες επιλογές παραμένουν αποθηκευμένες.", _
           vbCritical, "Αναπληρώσεις"

WorkflowExit:
    CleanupDecisionFiles
End Function

Private Function PrepareNextReplacementDecision() As Boolean
    Dim scriptPath As String, responsePath As String
    Dim commandLine As String, exitCode As Long

    CleanupDecisionFiles
    scriptPath = ResolveProjectTool("prepare_replacement_operator_decision.py")
    If Len(scriptPath) = 0 Then
        MsgBox "Η οθόνη ελεγχόμενων αναπληρώσεων δεν είναι διαθέσιμη.", _
               vbCritical, "Αναπληρώσεις"
        Exit Function
    End If

    mRowsPath = Environ$("TEMP") & "\eka_replacement_operator_rows_" & _
                Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\eka_replacement_operator_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --rows " & QuoteArg(mRowsPath) & _
                  " --response " & QuoteArg(responsePath)

    Application.StatusBar = "Προετοιμασία επόμενης απόφασης αναπλήρωσης..."
    exitCode = RunHidden(commandLine)
    Application.StatusBar = False

    If Dir$(responsePath) <> "" Then mDecisionResponse = ReadUtf8Text(responsePath)
    On Error Resume Next
    If Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0

    If exitCode <> 0 Or _
       InStr(1, mDecisionResponse, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or _
       Dir$(mRowsPath) = "" Then
        WriteOperatorDiagnostic exitCode, mDecisionResponse, "prepare decision"
        MsgBox "Δεν ήταν δυνατή η προετοιμασία των επιλογών αναπλήρωσης.", _
               vbExclamation, "Αναπληρώσεις"
        Exit Function
    End If

    PrepareNextReplacementDecision = True
End Function

Public Function ConfirmReplacementOperatorChoice( _
    ByVal targetSessionId As String, _
    ByVal providerId As String, _
    ByVal replacementTime As String) As Boolean

    Dim scriptPath As String, responsePath As String
    Dim commandLine As String, responseText As String
    Dim exitCode As Long, writeRow As Long
    Dim factSheetCreated As Boolean

    On Error GoTo ConfirmError

    scriptPath = ResolveProjectTool("validate_replacement_operator_choice.py")
    If Len(scriptPath) = 0 Then
        MsgBox "Η επαλήθευση της επιλογής δεν είναι διαθέσιμη.", vbCritical, "Αναπληρώσεις"
        Exit Function
    End If

    responsePath = Environ$("TEMP") & "\eka_replacement_operator_choice_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --target-session-id " & QuoteArg(targetSessionId) & _
                  " --provider-id " & QuoteArg(providerId) & _
                  " --time " & QuoteArg(replacementTime) & _
                  " --response " & QuoteArg(responsePath)

    Application.StatusBar = "Επαλήθευση επιλογής αναπλήρωσης..."
    exitCode = RunHidden(commandLine)
    Application.StatusBar = False
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)

    If exitCode <> 0 Or _
       InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η επιλογή δεν είναι πλέον διαθέσιμη. Η λίστα θα υπολογιστεί ξανά.", _
               vbExclamation, "Αναπληρώσεις"
        GoTo ConfirmExit
    End If

    If Not PersistValidatedOperatorReplacement(responseText, writeRow, factSheetCreated) Then _
        GoTo ConfirmExit

    If Not ApplyOperatorProjectionRefresh() Then
        RollbackOperatorReplacement writeRow, factSheetCreated
        MsgBox "Η επιλογή δεν εφαρμόστηκε και η εγγραφή αναπλήρωσης αναιρέθηκε.", _
               vbCritical, "Αναπληρώσεις"
        GoTo ConfirmExit
    End If

    mDecisionAccepted = True
    ConfirmReplacementOperatorChoice = True

ConfirmExit:
    Application.StatusBar = False
    On Error Resume Next
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
    Exit Function

ConfirmError:
    Application.StatusBar = False
    WriteOperatorDiagnostic Err.Number, Err.Description, "confirm choice"
    If writeRow > 0 Then RollbackOperatorReplacement writeRow, factSheetCreated
    MsgBox "Η επιλογή αναπλήρωσης δεν αποθηκεύτηκε.", vbCritical, "Αναπληρώσεις"
    Resume ConfirmExit
End Function

Public Sub CancelReplacementOperatorWorkflow()
    mWorkflowCancelled = True
End Sub

Private Function PersistValidatedOperatorReplacement( _
    ByVal responseText As String, _
    ByRef writeRow As Long, _
    ByRef sheetCreated As Boolean) As Boolean

    Dim ws As Worksheet, lastRow As Long, row As Long
    Dim replacementId As String, targetId As String, patientId As String
    Dim originalTherapist As String, replacementTherapist As String
    Dim replacementDate As String, replacementTime As String
    Dim reason As String, providerKind As String, statusText As String

    replacementId = ReplacementJsonStringValue(responseText, "replacement_id")
    targetId = ReplacementJsonStringValue(responseText, "target_session_id")
    patientId = ReplacementJsonStringValue(responseText, "patient_id")
    originalTherapist = ReplacementJsonStringValue(responseText, "original_therapist_id")
    replacementTherapist = ReplacementJsonStringValue(responseText, "replacement_therapist_id")
    replacementDate = ReplacementJsonStringValue(responseText, "replacement_date")
    replacementTime = ReplacementJsonStringValue(responseText, "replacement_time")
    reason = ReplacementJsonStringValue(responseText, "reason")
    providerKind = ReplacementJsonStringValue(responseText, "provider_kind")
    statusText = ReplacementJsonStringValue(responseText, "status")

    If Len(replacementId) = 0 Or Len(targetId) = 0 Or Len(patientId) = 0 Or _
       Len(originalTherapist) = 0 Or Len(replacementTherapist) = 0 Or _
       Len(replacementDate) = 0 Or Len(replacementTime) = 0 Or _
       UCase$(statusText) <> "ACTIVE" Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    On Error GoTo PersistError
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = FACT_SHEET
        WriteFactHeaders ws
        sheetCreated = True
    ElseIf Not FactHeadersAreValid(ws) Then
        Exit Function
    End If

    ws.Unprotect
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = replacementId Then GoTo PersistError
        If Trim$(CStr(ws.Cells(row, 2).Value)) = targetId And _
           UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = "ACTIVE" Then GoTo PersistError
    Next row

    writeRow = IIf(lastRow < 2, 2, lastRow + 1)
    ws.Cells(writeRow, 1).Value2 = replacementId
    ws.Cells(writeRow, 2).Value2 = targetId
    ws.Cells(writeRow, 3).Value2 = patientId
    ws.Cells(writeRow, 4).Value2 = originalTherapist
    ws.Cells(writeRow, 5).Value2 = replacementTherapist
    ws.Cells(writeRow, 6).NumberFormat = "@"
    ws.Cells(writeRow, 6).Value2 = replacementDate
    ws.Cells(writeRow, 7).NumberFormat = "@"
    ws.Cells(writeRow, 7).Value2 = replacementTime
    ws.Cells(writeRow, 8).Value2 = reason
    ws.Cells(writeRow, 9).Value2 = providerKind
    ws.Cells(writeRow, 10).Value2 = "ACTIVE"

    If Trim$(CStr(ws.Cells(writeRow, 1).Value)) <> replacementId Or _
       Trim$(CStr(ws.Cells(writeRow, 2).Value)) <> targetId Or _
       Trim$(CStr(ws.Cells(writeRow, 5).Value)) <> replacementTherapist Or _
       Trim$(CStr(ws.Cells(writeRow, 7).Value)) <> replacementTime Or _
       UCase$(Trim$(CStr(ws.Cells(writeRow, 10).Value))) <> "ACTIVE" Then GoTo PersistError

    ws.Visible = xlSheetVeryHidden
    ws.Protect
    ThisWorkbook.Save
    PersistValidatedOperatorReplacement = True
    Exit Function

PersistError:
    On Error Resume Next
    If Not ws Is Nothing Then
        If writeRow > 0 Then ws.Rows(writeRow).ClearContents
        ws.Visible = xlSheetVeryHidden
        ws.Protect
    End If
    On Error GoTo 0
End Function

Private Function ApplyOperatorProjectionRefresh() As Boolean
    Dim scriptPath As String, payloadPath As String, responsePath As String
    Dim commandLine As String, responseText As String
    Dim exitCode As Long

    scriptPath = ResolveProjectTool("prepare_daily_input_in_place.py")
    If Len(scriptPath) = 0 Then Exit Function

    payloadPath = Environ$("TEMP") & "\eka_operator_projection_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\eka_operator_projection_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(responsePath)

    exitCode = RunHidden(commandLine)
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or _
       InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or _
       Dir$(payloadPath) = "" Then GoTo RefreshExit

    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo RefreshExit
    On Error GoTo RefreshExit
    RenderTherapistWorkloadStrip
    ApplyMasterView
    KeepApplicationShell
    ThisWorkbook.Save
    ApplyOperatorProjectionRefresh = True

RefreshExit:
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
End Function

Private Sub RollbackOperatorReplacement(ByVal writeRow As Long, ByVal sheetCreated As Boolean)
    Dim ws As Worksheet, previousAlerts As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    If ws Is Nothing Then Exit Sub
    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    If sheetCreated Then
        ws.Delete
    Else
        ws.Unprotect
        If writeRow > 0 Then ws.Rows(writeRow).ClearContents
        ws.Visible = xlSheetVeryHidden
        ws.Protect
    End If
    Application.DisplayAlerts = previousAlerts
    ThisWorkbook.Save
    On Error GoTo 0
End Sub

Private Sub CleanupDecisionFiles()
    On Error Resume Next
    If Len(mRowsPath) > 0 And Dir$(mRowsPath) <> "" Then Kill mRowsPath
    mRowsPath = ""
    mDecisionResponse = ""
    On Error GoTo 0
End Sub

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

Private Function RunHidden(ByVal commandLine As String) As Long
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    RunHidden = shell.Run(commandLine, 0, True)
End Function

Private Function ResolveProjectTool(ByVal fileName As String) As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\" & fileName
    If Dir$(candidate) <> "" Then ResolveProjectTool = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\" & fileName
    If Dir$(candidate) <> "" Then _
        ResolveProjectTool = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
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

Public Function ReplacementJsonStringValue(ByVal jsonText As String, ByVal key As String) As String
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
    ReplacementJsonStringValue = value
End Function

Public Function ReplacementJsonLongValue(ByVal jsonText As String, ByVal key As String) As Long
    Dim marker As String, startPos As Long, endPos As Long, rawValue As String
    marker = Chr$(34) & key & Chr$(34) & ":"
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(marker)
    Do While startPos <= Len(jsonText) And Mid$(jsonText, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    endPos = startPos
    Do While endPos <= Len(jsonText) And Mid$(jsonText, endPos, 1) Like "[0-9]"
        endPos = endPos + 1
    Loop
    rawValue = Mid$(jsonText, startPos, endPos - startPos)
    If Len(rawValue) > 0 Then ReplacementJsonLongValue = CLng(rawValue)
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Sub WriteOperatorDiagnostic(ByVal errorNumber As Long, ByVal errorText As String, ByVal phase As String)
    Dim filePath As String, handle As Integer
    filePath = Environ$("TEMP") & "\eka_replacement_operator_last_error.log"
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #handle, "phase=" & phase
    Print #handle, "error_number=" & CStr(errorNumber)
    Print #handle, "error=" & errorText
    Close #handle
End Sub
'''

# VBA uses a literal backslash; Python raw strings above use doubled backslashes for
# readability in source. Normalize before installing, matching the DAILY_INPUT bridge.
MODULE_CODE = MODULE_CODE.replace("\\\\", "\\")


FORM_CODE = r'''Option Explicit

Private mResponseText As String

Private Sub UserForm_Initialize()
    Me.Caption = "Τακτοποίηση αναπληρώσεων"
    Me.Width = 690
    Me.Height = 455
    Me.StartUpPosition = 1
    Me.BackColor = RGB(239, 249, 246)

    lblTitle.Caption = "Επιλογή αναπλήρωσης"
    lblTitle.Font.Name = "Calibri"
    lblTitle.Font.Size = 16
    lblTitle.Font.Bold = True
    lblTitle.ForeColor = RGB(38, 86, 76)

    lblProgress.Font.Name = "Calibri"
    lblProgress.Font.Size = 10
    lblProgress.ForeColor = RGB(90, 105, 100)

    lblPatient.Font.Name = "Calibri"
    lblPatient.Font.Size = 14
    lblPatient.Font.Bold = True
    lblPatient.ForeColor = RGB(55, 64, 72)

    lblOriginal.Font.Name = "Calibri"
    lblOriginal.Font.Size = 11
    lblOriginal.ForeColor = RGB(85, 92, 98)

    lblHint.Caption = "Επίλεξε θεραπευτή και ώρα ή χρησιμοποίησε την καλύτερη αυτόματη πρόταση. Τίποτα δεν αποθηκεύεται χωρίς Επιβεβαίωση."
    lblHint.WordWrap = True
    lblHint.Font.Name = "Calibri"
    lblHint.Font.Size = 10
    lblHint.ForeColor = RGB(94, 69, 73)

    lstOptions.ColumnCount = 10
    lstOptions.ColumnWidths = "0 pt;0 pt;0 pt;0 pt;0 pt;175 pt;65 pt;60 pt;240 pt;0 pt"
    lstOptions.Font.Name = "Calibri"
    lstOptions.Font.Size = 10

    lblSelected.Caption = ""
    lblSelected.Font.Name = "Calibri"
    lblSelected.Font.Size = 10
    lblSelected.Font.Bold = True
    lblSelected.ForeColor = RGB(101, 50, 56)

    cmdAuto.Caption = "Καλύτερη αυτόματη πρόταση"
    cmdAuto.Font.Bold = True
    cmdConfirm.Caption = "Επιβεβαίωση επιλογής"
    cmdConfirm.Font.Bold = True
    cmdCancel.Caption = "Διακοπή"
    cmdCancel.Cancel = True
End Sub

Public Sub LoadDecision(ByVal rowsPath As String, ByVal responseText As String)
    Dim stream As Object, text As String, lines() As String
    Dim fields() As String, i As Long
    Dim pending As Long

    mResponseText = responseText
    pending = modReplacementOperator.ReplacementJsonLongValue(responseText, "pending_count")
    lblProgress.Caption = "Εκκρεμείς συνεδρίες: " & CStr(pending)
    lblPatient.Caption = modReplacementOperator.ReplacementJsonStringValue(responseText, "patient_name")
    lblOriginal.Caption = "Αρχικό πρόγραμμα: " & _
                          modReplacementOperator.ReplacementJsonStringValue(responseText, "original_therapist") & _
                          "  |  " & modReplacementOperator.ReplacementJsonStringValue(responseText, "original_time")

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile rowsPath
    text = stream.ReadText
    stream.Close

    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    lstOptions.Clear
    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            fields = Split(lines(i), vbTab)
            If UBound(fields) >= 9 Then
                lstOptions.AddItem fields(0)
                lstOptions.List(lstOptions.ListCount - 1, 1) = fields(1)
                lstOptions.List(lstOptions.ListCount - 1, 2) = fields(2)
                lstOptions.List(lstOptions.ListCount - 1, 3) = fields(3)
                lstOptions.List(lstOptions.ListCount - 1, 4) = fields(4)
                lstOptions.List(lstOptions.ListCount - 1, 5) = fields(5)
                lstOptions.List(lstOptions.ListCount - 1, 6) = fields(6)
                lstOptions.List(lstOptions.ListCount - 1, 7) = fields(7)
                lstOptions.List(lstOptions.ListCount - 1, 8) = fields(8)
                lstOptions.List(lstOptions.ListCount - 1, 9) = fields(9)
            End If
        End If
    Next i

    If lstOptions.ListCount = 0 Then
        lblHint.Caption = "Δεν υπάρχει κανονική εφικτή λύση για αυτή τη συνεδρία. Δεν θα γίνει αυτόματη αναστολή ή διπλοκράτηση."
        cmdAuto.Enabled = False
        cmdConfirm.Enabled = False
    Else
        cmdAuto.Enabled = True
        cmdConfirm.Enabled = True
        lstOptions.ListIndex = 0
    End If
End Sub

Private Sub lstOptions_Change()
    Dim idx As Long
    If lstOptions.ListIndex < 0 Then
        lblSelected.Caption = ""
        Exit Sub
    End If
    idx = lstOptions.ListIndex
    lblSelected.Caption = "Επιλογή: " & lstOptions.List(idx, 5) & _
                          " στις " & lstOptions.List(idx, 6) & _
                          "  |  φόρτος " & lstOptions.List(idx, 7)
End Sub

Private Sub cmdAuto_Click()
    Dim i As Long
    For i = 0 To lstOptions.ListCount - 1
        If CStr(lstOptions.List(i, 9)) = "1" Then
            lstOptions.ListIndex = i
            lblSelected.Caption = "Αυτόματη πρόταση: " & lstOptions.List(i, 5) & _
                                  " στις " & lstOptions.List(i, 6) & _
                                  "  |  φόρτος " & lstOptions.List(i, 7)
            Exit Sub
        End If
    Next i
    MsgBox "Δεν υπάρχει διαθέσιμη αυτόματη πρόταση για αυτή τη συνεδρία.", _
           vbInformation, "Αναπληρώσεις"
End Sub

Private Sub cmdConfirm_Click()
    Dim idx As Long
    If lstOptions.ListIndex < 0 Then
        MsgBox "Επίλεξε θεραπευτή και ώρα.", vbExclamation, "Αναπληρώσεις"
        Exit Sub
    End If
    idx = lstOptions.ListIndex

    If MsgBox("Ασθενής: " & lstOptions.List(idx, 1) & vbCrLf & _
              "Αρχικό: " & lstOptions.List(idx, 2) & " " & lstOptions.List(idx, 3) & vbCrLf & _
              "Νέα ανάθεση: " & lstOptions.List(idx, 5) & " " & lstOptions.List(idx, 6) & vbCrLf & vbCrLf & _
              "Να αποθηκευτεί αυτή η επιλογή;", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Επιβεβαίωση αναπλήρωσης") <> vbYes Then
        Exit Sub
    End If

    If modReplacementOperator.ConfirmReplacementOperatorChoice( _
        CStr(lstOptions.List(idx, 0)), _
        CStr(lstOptions.List(idx, 4)), _
        CStr(lstOptions.List(idx, 6))) Then
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    modReplacementOperator.CancelReplacementOperatorWorkflow
    Unload Me
End Sub
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _patch_daily_input_route(vbproject) -> None:
    component = vbproject.VBComponents(DAILY_INPUT_MODULE_NAME)
    code_module = component.CodeModule
    source = str(code_module.Lines(1, code_module.CountOfLines))
    legacy = "If ResolveDailyReplacementsInPlace() Then"
    routed = "If RunReplacementOperatorWorkflow() Then"
    if routed in source:
        return
    if legacy not in source:
        raise RuntimeError("DAILY_INPUT replacement continuation point was not found")
    source = source.replace(legacy, routed, 1)
    code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(source)


def install_replacement_operator(vbproject) -> None:
    _remove_component_if_present(vbproject, MODULE_NAME)
    _remove_component_if_present(vbproject, FORM_NAME)

    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)

    form = vbproject.VBComponents.Add(3)
    form.Name = FORM_NAME
    designer = form.Designer
    designer.Caption = "Τακτοποίηση αναπληρώσεων"

    controls = [
        ("Forms.Label.1", "lblTitle", 22, 16, 620, 28),
        ("Forms.Label.1", "lblProgress", 22, 48, 620, 18),
        ("Forms.Label.1", "lblPatient", 22, 72, 620, 26),
        ("Forms.Label.1", "lblOriginal", 22, 102, 620, 22),
        ("Forms.Label.1", "lblHint", 22, 132, 620, 42),
        ("Forms.ListBox.1", "lstOptions", 22, 182, 620, 150),
        ("Forms.Label.1", "lblSelected", 22, 340, 620, 22),
        ("Forms.CommandButton.1", "cmdAuto", 22, 374, 205, 34),
        ("Forms.CommandButton.1", "cmdConfirm", 242, 374, 205, 34),
        ("Forms.CommandButton.1", "cmdCancel", 462, 374, 180, 34),
    ]
    for prog_id, name, left, top, width, height in controls:
        control = designer.Controls.Add(prog_id, name, True)
        control.Left = left
        control.Top = top
        control.Width = width
        control.Height = height

    form.CodeModule.AddFromString(FORM_CODE)
    _patch_daily_input_route(vbproject)
