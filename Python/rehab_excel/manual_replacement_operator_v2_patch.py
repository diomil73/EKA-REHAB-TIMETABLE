from __future__ import annotations

from pathlib import Path
import sys

from rehab_excel.replacement_operator_vba import MODULE_CODE as BASE_MODULE_CODE


MODULE_NAME = "modReplacementOperator"
FORM_NAME = "frmReplacementDecision"
LEDGER_MODULE_NAME = "modDailyChangeLedger"
PATCH_MARKER = "MANUAL_OPERATOR_V2"


class ManualReplacementOperatorV2PatchError(RuntimeError):
    pass


def _replace_once(source: str, old: str, new: str, label: str) -> str:
    if old not in source:
        raise ManualReplacementOperatorV2PatchError(f"{label} marker not found")
    return source.replace(old, new, 1)


def build_v2_module_code(source: str = BASE_MODULE_CODE) -> str:
    if PATCH_MARKER in source:
        return source

    source = _replace_once(
        source,
        "Private mWorkflowCancelled As Boolean\n",
        "Private mWorkflowCancelled As Boolean\n"
        "Private mAdvanceRequested As Boolean\n"
        "Private mPreviousRequested As Boolean\n"
        "Private mReloadRequested As Boolean\n"
        "Private mAutomaticRemainingRequested As Boolean\n"
        "Private mCurrentDecisionKind As String\n"
        "Private mCurrentDecisionToken As String\n"
        "Private mDecisionHistoryKinds As Collection\n"
        "Private mDecisionHistoryTokens As Collection\n"
        f"Private Const {PATCH_MARKER} As String = \"{PATCH_MARKER}\"\n",
        "module declarations",
    )

    old_reset = '''        mDecisionAccepted = False\n        mWorkflowCancelled = False\n        frmReplacementDecision.LoadDecision mRowsPath, mDecisionResponse\n        frmReplacementDecision.Show vbModal\n        If mWorkflowCancelled Then GoTo WorkflowExit\n        If Not mDecisionAccepted Then GoTo WorkflowExit\n'''
    new_reset = '''        If mDecisionHistoryKinds Is Nothing Then Set mDecisionHistoryKinds = New Collection\n        If mDecisionHistoryTokens Is Nothing Then Set mDecisionHistoryTokens = New Collection\n        mDecisionAccepted = False\n        mAdvanceRequested = False\n        mPreviousRequested = False\n        mReloadRequested = False\n        mAutomaticRemainingRequested = False\n        mCurrentDecisionKind = ""\n        mCurrentDecisionToken = ""\n        mWorkflowCancelled = False\n        frmReplacementDecision.LoadDecision mRowsPath, mDecisionResponse\n        frmReplacementDecision.Show vbModal\n        If mWorkflowCancelled Then GoTo WorkflowExit\n        If mAutomaticRemainingRequested Then\n            RunManualReplacementWorkflow = ResolveDailyReplacementsInPlace()\n            GoTo WorkflowExit\n        End If\n        If mPreviousRequested Or mReloadRequested Then GoTo ContinueWorkflow\n        If Not mAdvanceRequested Then GoTo WorkflowExit\n        If Not mDecisionAccepted Then GoTo WorkflowExit\nContinueWorkflow:\n'''
    source = _replace_once(source, old_reset, new_reset, "manual workflow state")

    old_success = '''    mDecisionAccepted = True\n    ConfirmReplacementOperatorChoice = True\n'''
    new_success = '''    mDecisionAccepted = True\n    mCurrentDecisionKind = "replacement"\n    mCurrentDecisionToken = ReplacementJsonStringValue(responseText, "replacement_id")\n    ConfirmReplacementOperatorChoice = True\n'''
    source = _replace_once(source, old_success, new_success, "replacement confirmation state")

    insertion_marker = '''Public Sub CancelReplacementOperatorWorkflow()\n    mWorkflowCancelled = True\nEnd Sub\n'''
    controls_block = r'''Public Function CurrentOperatorDecisionConfirmed() As Boolean
    CurrentOperatorDecisionConfirmed = mDecisionAccepted
End Function

Public Function ConfirmOperatorTreatmentSuspension(ByVal targetSessionId As String, ByVal patientName As String, ByVal originalTime As String) As Boolean
    Dim ws As Worksheet, headerRow As Long, targetRow As Long, row As Long
    On Error GoTo SuspensionError

    Set ws = ThisWorkbook.Worksheets("DAILY_INPUT")
    ws.Unprotect
    For row = 1 To 100
        If Trim$(CStr(ws.Cells(row, 1).Value)) = "Ασθενής" And _
           Trim$(CStr(ws.Cells(row, 4).Value)) = "Κατάσταση" Then
            headerRow = row
            Exit For
        End If
    Next row
    If headerRow = 0 Then Err.Raise vbObjectError + 2951, , "Δεν βρέθηκε η περιοχή ασθενών στο DAILY_INPUT."

    For row = headerRow + 1 To headerRow + 500
        If Len(Trim$(CStr(ws.Cells(row, 1).Value))) = 0 Then
            targetRow = row
            Exit For
        End If
    Next row
    If targetRow = 0 Then Err.Raise vbObjectError + 2952, , "Δεν υπάρχει κενή γραμμή για αναβολή θεραπείας."

    ws.Cells(targetRow, 1).Value2 = patientName
    ws.Cells(targetRow, 2).Value2 = "ΟΧΙ"
    ws.Cells(targetRow, 3).NumberFormat = "@"
    ws.Cells(targetRow, 3).Value2 = originalTime
    ws.Cells(targetRow, 4).Value2 = "Αναστολή θεραπείας"
    ws.Cells(targetRow, 5).Value2 = "Χειροκίνητη απόφαση αναπλήρωσης"
    ws.Protect UserInterfaceOnly:=True
    ThisWorkbook.Save

    If Not ApplyOperatorProjectionRefresh() Then
        ws.Unprotect
        ws.Range("A" & CStr(targetRow) & ":E" & CStr(targetRow)).ClearContents
        ws.Protect UserInterfaceOnly:=True
        ThisWorkbook.Save
        Exit Function
    End If

    mDecisionAccepted = True
    mCurrentDecisionKind = "suspension"
    mCurrentDecisionToken = CStr(targetRow)
    ConfirmOperatorTreatmentSuspension = True
    Exit Function

SuspensionError:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect UserInterfaceOnly:=True
    WriteOperatorDiagnostic Err.Number, Err.Description, "confirm treatment suspension"
End Function

Public Function RequestNextReplacementPatient() As Boolean
    If Not mDecisionAccepted Then
        MsgBox "Επιβεβαιώστε πρώτα την απόφαση για τον τρέχοντα ασθενή.", vbInformation, "Αναπληρώσεις"
        Exit Function
    End If
    If Len(mCurrentDecisionKind) > 0 And Len(mCurrentDecisionToken) > 0 Then
        If mDecisionHistoryKinds Is Nothing Then Set mDecisionHistoryKinds = New Collection
        If mDecisionHistoryTokens Is Nothing Then Set mDecisionHistoryTokens = New Collection
        mDecisionHistoryKinds.Add mCurrentDecisionKind
        mDecisionHistoryTokens.Add mCurrentDecisionToken
    End If
    mAdvanceRequested = True
    RequestNextReplacementPatient = True
End Function

Public Function RequestPreviousReplacementPatient() As Boolean
    Dim kind As String, token As String, index As Long
    If mDecisionAccepted Then
        MsgBox "Για επιστροφή στον προηγούμενο ασθενή, αναιρέστε πρώτα την τρέχουσα επιβεβαιωμένη απόφαση.", vbInformation, "Αναπληρώσεις"
        Exit Function
    End If
    If mDecisionHistoryKinds Is Nothing Or mDecisionHistoryKinds.Count = 0 Then
        MsgBox "Δεν υπάρχει προηγούμενη χειροκίνητη απόφαση.", vbInformation, "Αναπληρώσεις"
        Exit Function
    End If

    index = mDecisionHistoryKinds.Count
    kind = CStr(mDecisionHistoryKinds(index))
    token = CStr(mDecisionHistoryTokens(index))
    If Not UndoOperatorDecision(kind, token) Then Exit Function
    mDecisionHistoryKinds.Remove index
    mDecisionHistoryTokens.Remove index
    mPreviousRequested = True
    RequestPreviousReplacementPatient = True
End Function

Public Function RequestChangeCurrentOperatorDecision() As Boolean
    If Not mDecisionAccepted Then
        MsgBox "Η τρέχουσα απόφαση δεν έχει ακόμη επιβεβαιωθεί.", vbInformation, "Αναπληρώσεις"
        Exit Function
    End If
    If Not UndoOperatorDecision(mCurrentDecisionKind, mCurrentDecisionToken) Then Exit Function
    mDecisionAccepted = False
    mCurrentDecisionKind = ""
    mCurrentDecisionToken = ""
    mReloadRequested = True
    RequestChangeCurrentOperatorDecision = True
End Function

Public Sub RequestAutomaticRemainingReplacements()
    mAutomaticRemainingRequested = True
End Sub

Private Function UndoOperatorDecision(ByVal kind As String, ByVal token As String) As Boolean
    Dim ws As Worksheet
    On Error GoTo UndoError
    If kind = "replacement" Then
        UndoOperatorDecision = modDailyChangeLedger.RestoreReplacementForOperator(token)
        Exit Function
    End If
    If kind = "suspension" Then
        Set ws = ThisWorkbook.Worksheets("DAILY_INPUT")
        ws.Unprotect
        ws.Range("A" & token & ":E" & token).ClearContents
        ws.Protect UserInterfaceOnly:=True
        ThisWorkbook.Save
        UndoOperatorDecision = ApplyOperatorProjectionRefresh()
        Exit Function
    End If
    Exit Function
UndoError:
    WriteOperatorDiagnostic Err.Number, Err.Description, "undo operator decision"
End Function

Public Sub CancelReplacementOperatorWorkflow()
    mWorkflowCancelled = True
End Sub
'''
    source = _replace_once(source, insertion_marker, controls_block, "operator control insertion")

    old_refresh_start = '''    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then GoTo RefreshExit\n    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo RefreshExit\n'''
    new_refresh_start = '''    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then GoTo RefreshExit\n    On Error Resume Next\n    ThisWorkbook.Worksheets("THERAPIST_DAILY").Unprotect\n    On Error GoTo 0\n    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo RefreshExit\n'''
    source = _replace_once(source, old_refresh_start, new_refresh_start, "projection unprotect")

    old_refresh_exit = '''RefreshExit:\n    On Error Resume Next\n    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath\n'''
    new_refresh_exit = '''RefreshExit:\n    On Error Resume Next\n    ThisWorkbook.Worksheets("THERAPIST_DAILY").Protect UserInterfaceOnly:=True\n    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath\n'''
    source = _replace_once(source, old_refresh_exit, new_refresh_exit, "projection reprotect")
    return source


def patch_ledger_module_for_operator(source: str) -> str:
    if "Public Function RestoreReplacementForOperator" in source:
        return source
    marker = "Public Sub RestoreDailyChangeRow()\n"
    wrapper = r'''Public Function RestoreReplacementForOperator(ByVal replacementId As String) As Boolean
    If Len(Trim$(replacementId)) = 0 Then Exit Function
    RestoreReplacementForOperator = RunAuthoritativeRestore(" --replacement-id " & QuoteArg(replacementId), False)
End Function

Public Sub RestoreDailyChangeRow()
'''
    return _replace_once(source, marker, wrapper, "ledger operator restore wrapper")


V2_FORM_CODE = r'''Option Explicit

Private mSelectedRow As Long
Private mTargetSessionId As String
Private mPatientName As String
Private mOriginalTime As String

Private Sub UserForm_Initialize()
    Me.Caption = "Χειροκίνητη τακτοποίηση αναπληρώσεων"
    Me.Width = 760: Me.Height = 500: Me.StartUpPosition = 1
    lblTitle.Caption = "Επιλογή αναπλήρωσης"
    lblTitle.Font.Size = 16: lblTitle.Font.Bold = True
    lblHint.Caption = "Επίλεξε θεραπευτή και μετά συγκεκριμένη ώρα. Η επιβεβαίωση δεν αλλάζει ασθενή."
    lblHint.WordWrap = True
    lstOptions.Visible = False
    cmbProvider.ColumnCount = 2
    cmbProvider.ColumnWidths = "180 pt;0 pt"
    cmdSuspend.Caption = "Αναβολή θεραπείας για σήμερα"
    cmdPrevious.Caption = "< Προηγούμενος ασθενής"
    cmdNext.Caption = "Επόμενος ασθενής >"
    cmdAutoRemaining.Caption = "Αυτόματη τακτοποίηση υπολοίπων"
    cmdChange.Caption = "Αλλαγή τρέχουσας επιλογής"
    cmdConfirm.Caption = "Επιβεβαίωση επιλογής": cmdConfirm.Font.Bold = True
    cmdCancel.Caption = "Διακοπή": cmdCancel.Cancel = True
    lblStatus.Caption = "Κατάσταση: μη επιβεβαιωμένος"
End Sub

Public Sub LoadDecision(ByVal rowsPath As String, ByVal responseText As String)
    Dim stream As Object, text As String, lines() As String, fields() As String, i As Long
    Dim seen As Object, providerName As String, providerId As String
    mTargetSessionId = modReplacementOperator.ReplacementJsonStringValue(responseText, "target_session_id")
    mPatientName = modReplacementOperator.ReplacementJsonStringValue(responseText, "patient_name")
    mOriginalTime = modReplacementOperator.ReplacementJsonStringValue(responseText, "original_time")
    lblProgress.Caption = "Εκκρεμείς συνεδρίες: " & CStr(modReplacementOperator.ReplacementJsonLongValue(responseText, "pending_count"))
    lblPatient.Caption = mPatientName
    lblOriginal.Caption = "Αρχικό πρόγραμμα: " & modReplacementOperator.ReplacementJsonStringValue(responseText, "original_therapist") & "  |  " & mOriginalTime

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

    Set seen = CreateObject("Scripting.Dictionary")
    cmbProvider.Clear: cmbTime.Clear
    For i = 0 To lstOptions.ListCount - 1
        providerId = CStr(lstOptions.List(i, 4))
        providerName = CStr(lstOptions.List(i, 5))
        If Not seen.Exists(providerId) Then
            seen.Add providerId, True
            cmbProvider.AddItem providerName
            cmbProvider.List(cmbProvider.ListCount - 1, 1) = providerId
        End If
    Next i
    cmbProvider.ListIndex = -1
    cmbTime.ListIndex = -1
    mSelectedRow = -1
    cmdConfirm.Enabled = False
    cmdNext.Enabled = False
    cmdChange.Enabled = False
    lblSelected.Caption = "Επίλεξε θεραπευτή και ώρα ή αναβολή θεραπείας."
End Sub

Private Sub cmbProvider_Change()
    Dim i As Long, providerId As String
    cmbTime.Clear: mSelectedRow = -1: cmdConfirm.Enabled = False
    If cmbProvider.ListIndex < 0 Then Exit Sub
    providerId = CStr(cmbProvider.List(cmbProvider.ListIndex, 1))
    For i = 0 To lstOptions.ListCount - 1
        If CStr(lstOptions.List(i, 4)) = providerId Then cmbTime.AddItem CStr(lstOptions.List(i, 6))
    Next i
    cmbTime.ListIndex = -1
    lblSelected.Caption = "Θεραπευτής: " & cmbProvider.Value & " | επίλεξε ώρα."
End Sub

Private Sub cmbTime_Change()
    Dim i As Long, providerId As String, selectedTime As String
    If cmbProvider.ListIndex < 0 Or cmbTime.ListIndex < 0 Then Exit Sub
    providerId = CStr(cmbProvider.List(cmbProvider.ListIndex, 1))
    selectedTime = CStr(cmbTime.Value)
    For i = 0 To lstOptions.ListCount - 1
        If CStr(lstOptions.List(i, 4)) = providerId And CStr(lstOptions.List(i, 6)) = selectedTime Then
            mSelectedRow = i
            lblSelected.Caption = "Επιλογή: " & CStr(lstOptions.List(i, 5)) & " στις " & selectedTime & " | φόρτος " & CStr(lstOptions.List(i, 7))
            cmdConfirm.Enabled = True
            Exit For
        End If
    Next i
End Sub

Private Sub cmdConfirm_Click()
    Dim idx As Long
    If mSelectedRow < 0 Then Exit Sub
    idx = mSelectedRow
    If MsgBox("Ασθενής: " & CStr(lstOptions.List(idx, 1)) & vbCrLf & _
              "Νέα ανάθεση: " & CStr(lstOptions.List(idx, 5)) & " " & CStr(lstOptions.List(idx, 6)) & vbCrLf & vbCrLf & _
              "Να αποθηκευτεί αυτή η επιλογή;", vbYesNo + vbQuestion + vbDefaultButton2, "Επιβεβαίωση αναπλήρωσης") <> vbYes Then Exit Sub
    lblStatus.Caption = "Κατάσταση: επαλήθευση και αποθήκευση..."
    Me.Repaint
    If modReplacementOperator.ConfirmReplacementOperatorChoice(CStr(lstOptions.List(idx, 0)), CStr(lstOptions.List(idx, 4)), CStr(lstOptions.List(idx, 6))) Then
        lblStatus.Caption = "Κατάσταση: ✓ επιβεβαιωμένος"
        cmdNext.Enabled = True: cmdChange.Enabled = True
        cmdConfirm.Enabled = False: cmdSuspend.Enabled = False
    Else
        lblStatus.Caption = "Κατάσταση: η επιλογή δεν αποθηκεύτηκε"
    End If
End Sub

Private Sub cmdSuspend_Click()
    If MsgBox("Να αναβληθεί η σημερινή θεραπεία του/της " & mPatientName & ";", vbYesNo + vbQuestion + vbDefaultButton2, "Αναβολή θεραπείας") <> vbYes Then Exit Sub
    lblStatus.Caption = "Κατάσταση: καταχώριση αναβολής...": Me.Repaint
    If modReplacementOperator.ConfirmOperatorTreatmentSuspension(mTargetSessionId, mPatientName, mOriginalTime) Then
        lblStatus.Caption = "Κατάσταση: ✓ επιβεβαιωμένη αναβολή"
        cmdNext.Enabled = True: cmdChange.Enabled = True
        cmdConfirm.Enabled = False: cmdSuspend.Enabled = False
    Else
        lblStatus.Caption = "Κατάσταση: η αναβολή δεν αποθηκεύτηκε"
    End If
End Sub

Private Sub cmdChange_Click()
    If modReplacementOperator.RequestChangeCurrentOperatorDecision() Then Unload Me
End Sub

Private Sub cmdNext_Click()
    If modReplacementOperator.RequestNextReplacementPatient() Then Unload Me
End Sub

Private Sub cmdPrevious_Click()
    If modReplacementOperator.RequestPreviousReplacementPatient() Then Unload Me
End Sub

Private Sub cmdAutoRemaining_Click()
    If MsgBox("Να γίνει αυτόματη τακτοποίηση μόνο των ασθενών που παραμένουν εκκρεμείς;", vbYesNo + vbQuestion + vbDefaultButton2, "Αυτόματη τακτοποίηση υπολοίπων") <> vbYes Then Exit Sub
    modReplacementOperator.RequestAutomaticRemainingReplacements
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    modReplacementOperator.CancelReplacementOperatorWorkflow
    Unload Me
End Sub
'''


def _ensure_control(designer, prog_id: str, name: str, left: int, top: int, width: int, height: int):
    try:
        control = designer.Controls(name)
    except Exception:
        control = designer.Controls.Add(prog_id, name, True)
    control.Left = left
    control.Top = top
    control.Width = width
    control.Height = height
    return control


def install_manual_operator_v2(workbook_path: str | Path) -> None:
    path = Path(workbook_path).resolve()
    if not path.exists():
        raise ManualReplacementOperatorV2PatchError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise ManualReplacementOperatorV2PatchError("Workbook must be an .xlsm file")
    if sys.platform != "win32":
        raise ManualReplacementOperatorV2PatchError("Windows Excel is required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ManualReplacementOperatorV2PatchError("pywin32 is required") from exc

    excel = None
    workbook = None
    failure: Exception | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject

        operator = vbproject.VBComponents(MODULE_NAME).CodeModule
        operator_source = str(operator.Lines(1, operator.CountOfLines))
        patched_operator = build_v2_module_code(operator_source)
        if patched_operator != operator_source:
            operator.DeleteLines(1, operator.CountOfLines)
            operator.AddFromString(patched_operator)

        ledger = vbproject.VBComponents(LEDGER_MODULE_NAME).CodeModule
        ledger_source = str(ledger.Lines(1, ledger.CountOfLines))
        patched_ledger = patch_ledger_module_for_operator(ledger_source)
        if patched_ledger != ledger_source:
            ledger.DeleteLines(1, ledger.CountOfLines)
            ledger.AddFromString(patched_ledger)

        form = vbproject.VBComponents(FORM_NAME)
        designer = form.Designer
        _ensure_control(designer, "Forms.ComboBox.1", "cmbProvider", 22, 178, 250, 24)
        _ensure_control(designer, "Forms.ComboBox.1", "cmbTime", 292, 178, 110, 24)
        _ensure_control(designer, "Forms.CommandButton.1", "cmdSuspend", 422, 176, 220, 28)
        _ensure_control(designer, "Forms.Label.1", "lblStatus", 22, 214, 620, 22)
        _ensure_control(designer, "Forms.CommandButton.1", "cmdPrevious", 22, 400, 155, 32)
        _ensure_control(designer, "Forms.CommandButton.1", "cmdChange", 187, 400, 165, 32)
        _ensure_control(designer, "Forms.CommandButton.1", "cmdNext", 362, 400, 140, 32)
        _ensure_control(designer, "Forms.CommandButton.1", "cmdAutoRemaining", 512, 400, 190, 32)

        try:
            designer.Controls("lstOptions").Visible = False
            designer.Controls("lblSelected").Top = 244
            designer.Controls("cmdConfirm").Left = 235
            designer.Controls("cmdConfirm").Top = 350
            designer.Controls("cmdCancel").Left = 462
            designer.Controls("cmdCancel").Top = 350
        except Exception:
            pass

        form_code = form.CodeModule
        form_code.DeleteLines(1, form_code.CountOfLines)
        form_code.AddFromString(V2_FORM_CODE)

        workbook.Save()
    except Exception as exc:
        failure = exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    if failure is not None:
        raise ManualReplacementOperatorV2PatchError(str(failure)) from failure
