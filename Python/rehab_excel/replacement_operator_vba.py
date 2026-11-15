from __future__ import annotations

MODULE_NAME = "modReplacementOperator"
FORM_NAME = "frmReplacementDecision"
MACRO_NAME = "ResolveReplacementWorkflow"
DAILY_INPUT_MODULE_NAME = "modDailyInputAction"

MODULE_CODE = r'''Option Explicit

Private Const FACT_SHEET As String = "DAILY_REPLACEMENT_FACTS"
Private mRowsPath As String
Private mDecisionResponse As String
Private mDecisionAccepted As Boolean
Private mWorkflowCancelled As Boolean

Public Function ResolveReplacementWorkflow() As Boolean
    Dim answer As VbMsgBoxResult

    answer = MsgBox( _
        "Υπάρχουν ασθενείς που χρειάζονται αναπλήρωση." & vbCrLf & vbCrLf & _
        "ΝΑΙ: καλύτερη αυτόματη συνολική λύση" & vbCrLf & _
        "ΟΧΙ: χειροκίνητη τακτοποίηση ασθενή-ασθενή" & vbCrLf & _
        "ΑΚΥΡΟ: διακοπή χωρίς νέα αναπλήρωση", _
        vbYesNoCancel + vbQuestion + vbDefaultButton2, _
        "Τακτοποίηση αναπληρώσεων")

    Select Case answer
        Case vbYes
            ResolveReplacementWorkflow = ResolveDailyReplacementsInPlace()
        Case vbNo
            ResolveReplacementWorkflow = RunManualReplacementWorkflow()
        Case Else
            ResolveReplacementWorkflow = False
    End Select
End Function

Public Function RunManualReplacementWorkflow() As Boolean
    Dim iteration As Long, mode As String
    On Error GoTo WorkflowError

    For iteration = 1 To 100
        If Not PrepareNextReplacementDecision() Then GoTo WorkflowExit
        mode = ReplacementJsonStringValue(mDecisionResponse, "mode")
        If mode = "complete" Then
            RunManualReplacementWorkflow = True
            GoTo WorkflowExit
        End If
        If mode <> "decision_required" Then
            MsgBox "Η διαδικασία αναπλήρωσης επέστρεψε μη αναμενόμενη κατάσταση.", vbCritical, "Αναπληρώσεις"
            GoTo WorkflowExit
        End If

        mDecisionAccepted = False
        mWorkflowCancelled = False
        frmReplacementDecision.LoadDecision mRowsPath, mDecisionResponse
        frmReplacementDecision.Show vbModal
        If mWorkflowCancelled Then GoTo WorkflowExit
        If Not mDecisionAccepted Then GoTo WorkflowExit
    Next iteration

    MsgBox "Η διαδικασία σταμάτησε για λόγους ασφαλείας μετά από πολλές διαδοχικές αποφάσεις.", vbExclamation, "Αναπληρώσεις"
    GoTo WorkflowExit

WorkflowError:
    WriteOperatorDiagnostic Err.Number, Err.Description, "manual workflow"
    MsgBox "Η χειροκίνητη τακτοποίηση δεν ολοκληρώθηκε. Οι ήδη επιβεβαιωμένες επιλογές παραμένουν αποθηκευμένες.", vbCritical, "Αναπληρώσεις"
WorkflowExit:
    CleanupDecisionFiles
End Function

Private Function PrepareNextReplacementDecision() As Boolean
    Dim scriptPath As String, responsePath As String, commandLine As String, exitCode As Long
    CleanupDecisionFiles
    scriptPath = ResolveProjectTool("prepare_replacement_operator_decision.py")
    If Len(scriptPath) = 0 Then
        MsgBox "Η χειροκίνητη τακτοποίηση δεν είναι διαθέσιμη.", vbCritical, "Αναπληρώσεις"
        Exit Function
    End If

    mRowsPath = Environ$("TEMP") & "\eka_replacement_operator_rows_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\eka_replacement_operator_response_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --rows " & QuoteArg(mRowsPath) & _
                  " --response " & QuoteArg(responsePath)
    Application.StatusBar = "Υπολογισμός διαθέσιμων επιλογών..."
    exitCode = RunHidden(commandLine)
    Application.StatusBar = False
    If Dir$(responsePath) <> "" Then mDecisionResponse = ReadUtf8Text(responsePath)
    On Error Resume Next
    If Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0

    If exitCode <> 0 Or InStr(1, mDecisionResponse, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(mRowsPath) = "" Then
        WriteOperatorDiagnostic exitCode, mDecisionResponse, "prepare decision"
        MsgBox "Δεν ήταν δυνατός ο υπολογισμός των επιλογών αναπλήρωσης.", vbExclamation, "Αναπληρώσεις"
        Exit Function
    End If
    PrepareNextReplacementDecision = True
End Function

Public Function ConfirmReplacementOperatorChoice(ByVal targetSessionId As String, ByVal providerId As String, ByVal replacementTime As String) As Boolean
    Dim scriptPath As String, responsePath As String, commandLine As String, responseText As String
    Dim exitCode As Long, writeRow As Long, factSheetCreated As Boolean
    On Error GoTo ConfirmError

    scriptPath = ResolveProjectTool("validate_replacement_operator_choice.py")
    If Len(scriptPath) = 0 Then Exit Function
    responsePath = Environ$("TEMP") & "\eka_replacement_operator_choice_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --target-session-id " & QuoteArg(targetSessionId) & _
                  " --provider-id " & QuoteArg(providerId) & _
                  " --time " & QuoteArg(replacementTime) & _
                  " --response " & QuoteArg(responsePath)
    Application.StatusBar = "Επαλήθευση επιλογής..."
    exitCode = RunHidden(commandLine)
    Application.StatusBar = False
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η επιλογή δεν είναι πλέον διαθέσιμη. Οι επιλογές θα υπολογιστούν ξανά.", vbExclamation, "Αναπληρώσεις"
        GoTo ConfirmExit
    End If

    If Not PersistValidatedOperatorReplacement(responseText, writeRow, factSheetCreated) Then GoTo ConfirmExit
    If Not ApplyOperatorProjectionRefresh() Then
        RollbackOperatorReplacement writeRow, factSheetCreated
        MsgBox "Η επιλογή δεν εφαρμόστηκε και αναιρέθηκε.", vbCritical, "Αναπληρώσεις"
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
    Resume ConfirmExit
End Function

Public Sub CancelReplacementOperatorWorkflow()
    mWorkflowCancelled = True
End Sub

Private Function PersistValidatedOperatorReplacement(ByVal responseText As String, ByRef writeRow As Long, ByRef sheetCreated As Boolean) As Boolean
    Dim ws As Worksheet, lastRow As Long, row As Long
    Dim replacementId As String, targetId As String, patientId As String
    Dim originalTherapist As String, replacementTherapist As String
    Dim replacementDate As String, replacementTime As String, reason As String, providerKind As String, statusText As String

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
    If Len(replacementId) = 0 Or Len(targetId) = 0 Or Len(patientId) = 0 Or Len(originalTherapist) = 0 Or Len(replacementTherapist) = 0 Or Len(replacementDate) = 0 Or Len(replacementTime) = 0 Or UCase$(statusText) <> "ACTIVE" Then Exit Function

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
        If Trim$(CStr(ws.Cells(row, 2).Value)) = targetId And UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = "ACTIVE" Then GoTo PersistError
    Next row
    writeRow = IIf(lastRow < 2, 2, lastRow + 1)
    ws.Cells(writeRow, 1).Value2 = replacementId
    ws.Cells(writeRow, 2).Value2 = targetId
    ws.Cells(writeRow, 3).Value2 = patientId
    ws.Cells(writeRow, 4).Value2 = originalTherapist
    ws.Cells(writeRow, 5).Value2 = replacementTherapist
    ws.Cells(writeRow, 6).NumberFormat = "@": ws.Cells(writeRow, 6).Value2 = replacementDate
    ws.Cells(writeRow, 7).NumberFormat = "@": ws.Cells(writeRow, 7).Value2 = replacementTime
    ws.Cells(writeRow, 8).Value2 = reason
    ws.Cells(writeRow, 9).Value2 = providerKind
    ws.Cells(writeRow, 10).Value2 = "ACTIVE"
    If Trim$(CStr(ws.Cells(writeRow, 1).Value)) <> replacementId Or Trim$(CStr(ws.Cells(writeRow, 2).Value)) <> targetId Or UCase$(Trim$(CStr(ws.Cells(writeRow, 10).Value))) <> "ACTIVE" Then GoTo PersistError
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
End Function

Private Function ApplyOperatorProjectionRefresh() As Boolean
    Dim scriptPath As String, payloadPath As String, responsePath As String, commandLine As String, responseText As String
    Dim exitCode As Long
    scriptPath = ResolveProjectTool("prepare_daily_input_in_place.py")
    If Len(scriptPath) = 0 Then Exit Function
    payloadPath = Environ$("TEMP") & "\eka_operator_projection_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\eka_operator_projection_response_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(responsePath)
    exitCode = RunHidden(commandLine)
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then GoTo RefreshExit
    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo RefreshExit
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
End Sub

Private Sub CleanupDecisionFiles()
    On Error Resume Next
    If Len(mRowsPath) > 0 And Dir$(mRowsPath) <> "" Then Kill mRowsPath
    mRowsPath = "": mDecisionResponse = ""
End Sub

Private Sub WriteFactHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value2 = "ReplacementId": ws.Cells(1, 2).Value2 = "TargetSessionId"
    ws.Cells(1, 3).Value2 = "PatientId": ws.Cells(1, 4).Value2 = "OriginalTherapistId"
    ws.Cells(1, 5).Value2 = "ReplacementTherapistId": ws.Cells(1, 6).Value2 = "ReplacementDate"
    ws.Cells(1, 7).Value2 = "ReplacementTime": ws.Cells(1, 8).Value2 = "Reason"
    ws.Cells(1, 9).Value2 = "ProviderKind": ws.Cells(1, 10).Value2 = "Status"
End Sub

Private Function FactHeadersAreValid(ByVal ws As Worksheet) As Boolean
    Dim headers As Variant, i As Long
    headers = Array("ReplacementId", "TargetSessionId", "PatientId", "OriginalTherapistId", "ReplacementTherapistId", "ReplacementDate", "ReplacementTime", "Reason", "ProviderKind", "Status")
    For i = 0 To UBound(headers)
        If Trim$(CStr(ws.Cells(1, i + 1).Value)) <> CStr(headers(i)) Then Exit Function
    Next i
    FactHeadersAreValid = True
End Function

Private Function RunHidden(ByVal commandLine As String) As Long
    RunHidden = CreateObject("WScript.Shell").Run(commandLine, 0, True)
End Function

Private Function ResolveProjectTool(ByVal fileName As String) As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\" & fileName
    If Dir$(candidate) <> "" Then ResolveProjectTool = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\" & fileName
    If Dir$(candidate) <> "" Then ResolveProjectTool = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
End Function

Private Function ReadUtf8Text(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.LoadFromFile filePath
    ReadUtf8Text = stream.ReadText
    stream.Close
End Function

Public Function ReplacementJsonStringValue(ByVal jsonText As String, ByVal key As String) As String
    Dim marker As String, startPos As Long, index As Long, ch As String, escaped As Boolean, value As String
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
    Do While startPos <= Len(jsonText) And Mid$(jsonText, startPos, 1) = " ": startPos = startPos + 1: Loop
    endPos = startPos
    Do While endPos <= Len(jsonText) And Mid$(jsonText, endPos, 1) Like "[0-9]": endPos = endPos + 1: Loop
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
MODULE_CODE = MODULE_CODE.replace("\\\\", "\\")

FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Χειροκίνητη τακτοποίηση αναπληρώσεων"
    Me.Width = 690: Me.Height = 430: Me.StartUpPosition = 1
    lblTitle.Caption = "Επιλογή αναπλήρωσης"
    lblTitle.Font.Size = 16: lblTitle.Font.Bold = True
    lblHint.Caption = "Επίλεξε θεραπευτή και ώρα. Μετά την επιβεβαίωση θα ξαναϋπολογιστούν οι επιλογές για τον επόμενο ασθενή."
    lblHint.WordWrap = True
    lstOptions.ColumnCount = 9
    lstOptions.ColumnWidths = "0 pt;0 pt;0 pt;0 pt;0 pt;180 pt;70 pt;70 pt;250 pt"
    cmdConfirm.Caption = "Επιβεβαίωση επιλογής": cmdConfirm.Font.Bold = True
    cmdCancel.Caption = "Διακοπή": cmdCancel.Cancel = True
End Sub

Public Sub LoadDecision(ByVal rowsPath As String, ByVal responseText As String)
    Dim stream As Object, text As String, lines() As String, fields() As String, i As Long
    lblProgress.Caption = "Εκκρεμείς συνεδρίες: " & CStr(modReplacementOperator.ReplacementJsonLongValue(responseText, "pending_count"))
    lblPatient.Caption = modReplacementOperator.ReplacementJsonStringValue(responseText, "patient_name")
    lblOriginal.Caption = "Αρχικό πρόγραμμα: " & modReplacementOperator.ReplacementJsonStringValue(responseText, "original_therapist") & "  |  " & modReplacementOperator.ReplacementJsonStringValue(responseText, "original_time")
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.LoadFromFile rowsPath
    text = stream.ReadText: stream.Close
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    lstOptions.Clear
    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            fields = Split(lines(i), vbTab)
            If UBound(fields) >= 8 Then
                lstOptions.AddItem fields(0)
                lstOptions.List(lstOptions.ListCount - 1, 1) = fields(1)
                lstOptions.List(lstOptions.ListCount - 1, 2) = fields(2)
                lstOptions.List(lstOptions.ListCount - 1, 3) = fields(3)
                lstOptions.List(lstOptions.ListCount - 1, 4) = fields(4)
                lstOptions.List(lstOptions.ListCount - 1, 5) = fields(5)
                lstOptions.List(lstOptions.ListCount - 1, 6) = fields(6)
                lstOptions.List(lstOptions.ListCount - 1, 7) = fields(7)
                lstOptions.List(lstOptions.ListCount - 1, 8) = fields(8)
            End If
        End If
    Next i
    If lstOptions.ListCount = 0 Then
        lblHint.Caption = "Δεν υπάρχει κανονική εφικτή επιλογή για αυτή τη συνεδρία. Διακόψτε και χρησιμοποιήστε την ελεγχόμενη διαδικασία εξαιρέσεων."
        cmdConfirm.Enabled = False
    Else
        cmdConfirm.Enabled = True
        lstOptions.ListIndex = 0
    End If
End Sub

Private Sub lstOptions_Change()
    If lstOptions.ListIndex < 0 Then lblSelected.Caption = "": Exit Sub
    lblSelected.Caption = "Επιλογή: " & lstOptions.List(lstOptions.ListIndex, 5) & " στις " & lstOptions.List(lstOptions.ListIndex, 6) & " | φόρτος " & lstOptions.List(lstOptions.ListIndex, 7)
End Sub

Private Sub cmdConfirm_Click()
    Dim idx As Long
    If lstOptions.ListIndex < 0 Then Exit Sub
    idx = lstOptions.ListIndex
    If MsgBox("Ασθενής: " & lstOptions.List(idx, 1) & vbCrLf & _
              "Αρχικό: " & lstOptions.List(idx, 2) & " " & lstOptions.List(idx, 3) & vbCrLf & _
              "Νέα ανάθεση: " & lstOptions.List(idx, 5) & " " & lstOptions.List(idx, 6) & vbCrLf & vbCrLf & _
              "Να αποθηκευτεί αυτή η επιλογή;", vbYesNo + vbQuestion + vbDefaultButton2, "Επιβεβαίωση αναπλήρωσης") <> vbYes Then Exit Sub
    If modReplacementOperator.ConfirmReplacementOperatorChoice(CStr(lstOptions.List(idx, 0)), CStr(lstOptions.List(idx, 4)), CStr(lstOptions.List(idx, 6))) Then Unload Me
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
    routed = "If ResolveReplacementWorkflow() Then"
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
    designer.Caption = "Χειροκίνητη τακτοποίηση αναπληρώσεων"
    controls = [
        ("Forms.Label.1", "lblTitle", 22, 16, 620, 28),
        ("Forms.Label.1", "lblProgress", 22, 48, 620, 18),
        ("Forms.Label.1", "lblPatient", 22, 72, 620, 26),
        ("Forms.Label.1", "lblOriginal", 22, 102, 620, 22),
        ("Forms.Label.1", "lblHint", 22, 132, 620, 42),
        ("Forms.ListBox.1", "lstOptions", 22, 182, 620, 140),
        ("Forms.Label.1", "lblSelected", 22, 330, 620, 22),
        ("Forms.CommandButton.1", "cmdConfirm", 235, 365, 205, 34),
        ("Forms.CommandButton.1", "cmdCancel", 462, 365, 180, 34),
    ]
    for prog_id, name, left, top, width, height in controls:
        control = designer.Controls.Add(prog_id, name, True)
        control.Left = left
        control.Top = top
        control.Width = width
        control.Height = height
    form.CodeModule.AddFromString(FORM_CODE)
    _patch_daily_input_route(vbproject)
