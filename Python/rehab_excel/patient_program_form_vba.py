from __future__ import annotations

PATIENT_PROGRAM_FORM_NAME = "frmPatientScheduling"
PATIENT_PROGRAM_MODULE_NAME = "PatientProgramRuntime"

PATIENT_PROGRAM_MODULE_CODE = r'''Option Explicit

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), "") & Chr$(34)
End Function

Private Function ProgramBridgePath() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\patient_program_bridge.py"
    If Len(Dir$(candidate)) > 0 Then ProgramBridgePath = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\Python\tools\patient_program_bridge.py"
    If Len(Dir$(candidate)) > 0 Then ProgramBridgePath = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\patient_program_bridge.py"
    If Len(Dir$(candidate)) > 0 Then ProgramBridgePath = candidate: Exit Function
    Err.Raise vbObjectError + 2910, "ProgramBridgePath", "Δεν βρέθηκε το patient_program_bridge.py."
End Function

Private Function PythonLauncher() As String
    Dim shell As Object
    Dim rc As Long
    Set shell = CreateObject("WScript.Shell")
    rc = shell.Run("cmd /d /c python --version >nul 2>&1", 0, True)
    If rc = 0 Then PythonLauncher = "python": Exit Function
    rc = shell.Run("cmd /d /c py -3 --version >nul 2>&1", 0, True)
    If rc = 0 Then PythonLauncher = "py -3": Exit Function
    Err.Raise vbObjectError + 2911, "PythonLauncher", "Δεν βρέθηκε διαθέσιμο Python runtime."
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

Private Function RunProgramBridge(ByVal arguments As String) As String
    Dim shell As Object
    Dim tempFile As String
    Dim command As String
    Dim rc As Long

    tempFile = Environ$("TEMP") & "\eka_patient_program_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    command = "cmd /d /c set PYTHONIOENCODING=utf-8&& " & PythonLauncher() & " " & _
              QuoteArg(ProgramBridgePath()) & " " & arguments & " > " & QuoteArg(tempFile) & " 2>&1"
    Set shell = CreateObject("WScript.Shell")
    rc = shell.Run(command, 0, True)
    RunProgramBridge = ReadUtf8Text(tempFile)
    On Error Resume Next
    Kill tempFile
    On Error GoTo 0
    If rc <> 0 And Left$(Trim$(RunProgramBridge), 5) <> "ERROR" Then _
        Err.Raise vbObjectError + 2912, "RunProgramBridge", Trim$(RunProgramBridge)
End Function

Public Function LoadPatientProgramConfig(ByVal patientId As String) As String
    LoadPatientProgramConfig = RunProgramBridge( _
        "config-vba " & QuoteArg(ThisWorkbook.FullName) & _
        " --patient-id " & QuoteArg(patientId) _
    )
End Function

Public Function FindPatientDisplayName(ByVal patientId As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Set ws = ThisWorkbook.Worksheets("PATIENTS")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 1).Value)), Trim$(patientId), vbTextCompare) = 0 Then
            FindPatientDisplayName = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
            Exit Function
        End If
    Next rowIndex
End Function

Public Sub OpenPatientScheduling(ByVal patientId As String)
    On Error GoTo OpenError
    Load frmPatientScheduling
    If frmPatientScheduling.Visible Then Exit Sub
    frmPatientScheduling.Configure patientId
    frmPatientScheduling.Show vbModeless
    Exit Sub
OpenError:
    Dim errorText As String
    errorText = Err.Description
    Err.Raise vbObjectError + 2913, "OpenPatientScheduling", errorText
End Sub
'''

PATIENT_PROGRAM_FORM_CODE = r'''Option Explicit

Private mPatientId As String
Private mCurrentTreatment As String
Private mLoading As Boolean
Private mProviderKinds As Object
Private mProviders As Object
Private mDraftDays As Object
Private mDraftTimes As Object
Private mDraftProviders As Object
Private mStandardSlots As Collection
Private mReclinedSlots As Collection

Public Sub Configure(ByVal patientId As String)
    Dim patientName As String
    mPatientId = Trim$(patientId)
    patientName = FindPatientDisplayName(mPatientId)
    If Len(patientName) > 0 Then
        lblPatient.Caption = "Ασθενής: " & patientName & "   |   Patient ID: " & mPatientId
    Else
        lblPatient.Caption = "Ασθενής: Patient ID " & mPatientId
    End If
    LoadProgramConfiguration
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Πρόγραμμα θεραπείας ασθενούς"
    Me.Width = 790
    Me.Height = 570
    Me.BackColor = RGB(245, 247, 250)

    lblTitle.Caption = "Πρόγραμμα θεραπείας ασθενούς"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 15
    lblHint.Caption = "Τσέκαρε όλες τις θεραπείες που χρειάζεται ο ασθενής. Μπορείς να συμπληρώσεις όσα στοιχεία θέλεις χειροκίνητα και να αφήσεις τα υπόλοιπα για αυτόματη συμπλήρωση."

    lstTreatments.MultiSelect = 1
    lstTreatments.ListStyle = 1
    cboDayPattern.Style = 2
    cboProvider.Style = 2
    cboTime.Style = 2

    Set mProviderKinds = CreateObject("Scripting.Dictionary")
    Set mProviders = CreateObject("Scripting.Dictionary")
    Set mDraftDays = CreateObject("Scripting.Dictionary")
    Set mDraftTimes = CreateObject("Scripting.Dictionary")
    Set mDraftProviders = CreateObject("Scripting.Dictionary")
    Set mStandardSlots = New Collection
    Set mReclinedSlots = New Collection

    cmdAutoComplete.Enabled = False
    cmdSaveProgram.Enabled = False
    lblStatus.Caption = "Πρώτα επίλεξε τις θεραπείες του ασθενούς."
End Sub

Private Sub LoadProgramConfiguration()
    Dim response As String
    Dim lines() As String
    Dim fields() As String
    Dim i As Long
    Dim treatment As String
    Dim provider As String

    On Error GoTo LoadError
    mLoading = True
    lstTreatments.Clear
    cboDayPattern.Clear
    cboProvider.Clear
    cboTime.Clear

    response = LoadPatientProgramConfig(mPatientId)
    lines = Split(Replace(response, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) = 0 Then GoTo NextLine
        fields = Split(lines(i), vbTab)
        If fields(0) = "ERROR" Then Err.Raise vbObjectError + 2920, , fields(1)
        Select Case fields(0)
            Case "TREATMENT"
                If UBound(fields) >= 4 Then
                    treatment = fields(2)
                    lstTreatments.AddItem treatment
                    mProviderKinds(treatment) = fields(3)
                    If Not mProviders.Exists(treatment) Then _
                        Set mProviders(treatment) = New Collection
                End If
            Case "PROVIDER"
                If UBound(fields) >= 2 Then
                    treatment = fields(1)
                    provider = fields(2)
                    If Not mProviders.Exists(treatment) Then _
                        Set mProviders(treatment) = New Collection
                    mProviders(treatment).Add provider
                End If
            Case "DAY_PATTERN"
                If UBound(fields) >= 1 Then cboDayPattern.AddItem fields(1)
            Case "STANDARD_SLOT"
                If UBound(fields) >= 1 Then mStandardSlots.Add fields(1)
            Case "RECLINED_SLOT"
                If UBound(fields) >= 1 Then mReclinedSlots.Add fields(1)
        End Select
NextLine:
    Next i

    mLoading = False
    lblStatus.Caption = CStr(lstTreatments.ListCount) & " θεραπείες διαθέσιμες από τις Ρυθμίσεις."
    Exit Sub
LoadError:
    mLoading = False
    lblStatus.Caption = "Δεν ήταν δυνατή η φόρτωση του προγράμματος: " & Err.Description
End Sub

Private Sub lstTreatments_Click()
    If mLoading Then Exit Sub
    SaveCurrentDraft
    If lstTreatments.ListIndex < 0 Then Exit Sub
    mCurrentTreatment = CStr(lstTreatments.List(lstTreatments.ListIndex))
    LoadCurrentTreatmentEditor
End Sub

Private Sub SaveCurrentDraft()
    If Len(mCurrentTreatment) = 0 Then Exit Sub
    If Len(Trim$(cboDayPattern.Value)) > 0 Then
        mDraftDays(mCurrentTreatment) = Trim$(cboDayPattern.Value)
    ElseIf mDraftDays.Exists(mCurrentTreatment) Then
        mDraftDays.Remove mCurrentTreatment
    End If
    If Len(Trim$(cboTime.Value)) > 0 Then
        mDraftTimes(mCurrentTreatment) = Trim$(cboTime.Value)
    ElseIf mDraftTimes.Exists(mCurrentTreatment) Then
        mDraftTimes.Remove mCurrentTreatment
    End If
    If Len(Trim$(cboProvider.Value)) > 0 Then
        mDraftProviders(mCurrentTreatment) = Trim$(cboProvider.Value)
    ElseIf mDraftProviders.Exists(mCurrentTreatment) Then
        mDraftProviders.Remove mCurrentTreatment
    End If
End Sub

Private Sub LoadCurrentTreatmentEditor()
    Dim item As Variant
    Dim treatment As String
    treatment = mCurrentTreatment
    mLoading = True

    lblEditing.Caption = "Ρυθμίσεις: " & treatment
    cboProvider.Clear
    cboTime.Clear

    If mProviders.Exists(treatment) Then
        For Each item In mProviders(treatment)
            cboProvider.AddItem CStr(item)
        Next item
    End If
    lblProvider.Visible = (cboProvider.ListCount > 0)
    cboProvider.Visible = (cboProvider.ListCount > 0)

    If StrComp(treatment, "Ανακλινόμενο", vbTextCompare) = 0 Then
        For Each item In mReclinedSlots
            cboTime.AddItem CStr(item)
        Next item
        lblPreference.Caption = "Προτίμηση: το Ανακλινόμενο θα τοποθετείται πριν ή μετά τη ΦΘ όπου είναι εφικτό."
    Else
        For Each item In mStandardSlots
            cboTime.AddItem CStr(item)
        Next item
        lblPreference.Caption = ""
    End If

    cboDayPattern.ListIndex = -1
    cboTime.ListIndex = -1
    cboProvider.ListIndex = -1
    If mDraftDays.Exists(treatment) Then cboDayPattern.Value = mDraftDays(treatment)
    If mDraftTimes.Exists(treatment) Then cboTime.Value = mDraftTimes(treatment)
    If mDraftProviders.Exists(treatment) Then cboProvider.Value = mDraftProviders(treatment)

    mLoading = False
    UpdateDraftStatus
End Sub

Private Sub cboDayPattern_Change()
    If Not mLoading Then SaveCurrentDraft: UpdateDraftStatus
End Sub

Private Sub cboTime_Change()
    If Not mLoading Then SaveCurrentDraft: UpdateDraftStatus
End Sub

Private Sub cboProvider_Change()
    If Not mLoading Then SaveCurrentDraft: UpdateDraftStatus
End Sub

Private Sub UpdateDraftStatus()
    Dim selectedCount As Long
    Dim i As Long
    For i = 0 To lstTreatments.ListCount - 1
        If lstTreatments.Selected(i) Then selectedCount = selectedCount + 1
    Next i
    cmdAutoComplete.Enabled = (selectedCount > 0)
    cmdSaveProgram.Enabled = False
    lblStatus.Caption = CStr(selectedCount) & " επιλεγμένες θεραπείες. Τα συμπληρωμένα πεδία θεωρούνται χειροκίνητες επιλογές."
End Sub

Private Sub cmdAutoComplete_Click()
    SaveCurrentDraft
    MsgBox "Η αυτόματη συμπλήρωση του νέου πολυθεραπευτικού draft είναι το επόμενο βήμα. Οι χειροκίνητες επιλογές που έβαλες παραμένουν κλειδωμένες.", vbInformation, "Αυτόματη συμπλήρωση"
End Sub

Private Sub cmdSaveProgram_Click()
    MsgBox "Η τελική atomic αποθήκευση θα ενεργοποιηθεί μαζί με τον νέο multi-treatment solver.", vbInformation, "Αποθήκευση προγράμματος"
End Sub

Private Sub cmdClose_Click()
    SaveCurrentDraft
    Unload Me
End Sub
'''


def _get_component(vbproject, component_name: str):
    try:
        return vbproject.VBComponents(component_name)
    except Exception:
        return None


def _replace_code(component, code: str) -> None:
    code_module = component.CodeModule
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    if count:
        code_module.DeleteLines(1, count)
    code_module.AddFromString(code)


def _ensure_control(designer, prog_id: str, name: str, caption: str | None, left: float, top: float, width: float, height: float, *, position_control):
    try:
        control = designer.Controls(name)
    except Exception:
        control = designer.Controls.Add(prog_id, name, True)
    if caption is not None:
        control.Caption = caption
    position_control(control, left, top)
    try:
        control.Width = width
        control.Height = height
    except Exception:
        pass
    return control


def install_patient_program_ui(vbproject, *, position_control) -> None:
    module = _get_component(vbproject, PATIENT_PROGRAM_MODULE_NAME)
    if module is None:
        module = vbproject.VBComponents.Add(1)
        module.Name = PATIENT_PROGRAM_MODULE_NAME
    _replace_code(module, PATIENT_PROGRAM_MODULE_CODE)

    form = _get_component(vbproject, PATIENT_PROGRAM_FORM_NAME)
    if form is None:
        form = vbproject.VBComponents.Add(3)
        form.Name = PATIENT_PROGRAM_FORM_NAME
    designer = form.Designer
    designer.Caption = "Πρόγραμμα θεραπείας ασθενούς"
    try:
        designer.Width = 760
        designer.Height = 540
    except Exception:
        pass

    controls = (
        ("Forms.Label.1", "lblTitle", "Πρόγραμμα θεραπείας ασθενούς", 18, 16, 520, 26),
        ("Forms.Label.1", "lblPatient", "Ασθενής:", 18, 48, 700, 22),
        ("Forms.Label.1", "lblHint", "", 18, 76, 710, 38),
        ("Forms.Label.1", "lblTreatments", "Θεραπείες", 18, 122, 250, 18),
        ("Forms.ListBox.1", "lstTreatments", None, 18, 144, 280, 300),
        ("Forms.Label.1", "lblEditing", "Ρυθμίσεις θεραπείας", 326, 122, 390, 20),
        ("Forms.Label.1", "lblDays", "Ημέρες", 326, 160, 130, 18),
        ("Forms.ComboBox.1", "cboDayPattern", None, 326, 180, 260, 22),
        ("Forms.Label.1", "lblTime", "Ώρα", 326, 218, 130, 18),
        ("Forms.ComboBox.1", "cboTime", None, 326, 238, 260, 22),
        ("Forms.Label.1", "lblProvider", "Θεραπευτής / provider", 326, 276, 180, 18),
        ("Forms.ComboBox.1", "cboProvider", None, 326, 296, 300, 22),
        ("Forms.Label.1", "lblPreference", "", 326, 336, 390, 42),
        ("Forms.Label.1", "lblStatus", "", 326, 388, 390, 40),
        ("Forms.CommandButton.1", "cmdAutoComplete", "Αυτόματη συμπλήρωση", 326, 446, 180, 30),
        ("Forms.CommandButton.1", "cmdSaveProgram", "Αποθήκευση προγράμματος", 516, 446, 200, 30),
        ("Forms.CommandButton.1", "cmdClose", "Κλείσιμο", 596, 486, 120, 30),
    )
    for prog_id, name, caption, left, top, width, height in controls:
        control = _ensure_control(
            designer, prog_id, name, caption, left, top, width, height,
            position_control=position_control,
        )
        if name in {"lblHint", "lblPreference", "lblStatus"}:
            try:
                control.WordWrap = True
            except Exception:
                pass

    _replace_code(form, PATIENT_PROGRAM_FORM_CODE)
