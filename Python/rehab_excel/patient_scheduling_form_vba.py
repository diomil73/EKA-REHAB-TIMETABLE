from __future__ import annotations

PATIENT_SCHEDULING_FORM_NAME = "frmPatientScheduling"
PATIENT_SCHEDULING_MODULE_NAME = "PatientSchedulingRuntime"

PATIENT_SCHEDULING_MODULE_CODE = r'''Option Explicit

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), "") & Chr$(34)
End Function

Private Function BridgePath() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\patient_scheduling_bridge.py"
    If Len(Dir$(candidate)) > 0 Then
        BridgePath = candidate
        Exit Function
    End If
    candidate = ThisWorkbook.Path & "\..\Python\tools\patient_scheduling_bridge.py"
    If Len(Dir$(candidate)) > 0 Then
        BridgePath = candidate
        Exit Function
    End If
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\patient_scheduling_bridge.py"
    If Len(Dir$(candidate)) > 0 Then
        BridgePath = candidate
        Exit Function
    End If
    Err.Raise vbObjectError + 2810, "BridgePath", "Δεν βρέθηκε το patient_scheduling_bridge.py."
End Function

Private Function PythonLauncher() As String
    Dim shell As Object
    Dim rc As Long
    Set shell = CreateObject("WScript.Shell")
    rc = shell.Run("cmd /d /c python --version >nul 2>&1", 0, True)
    If rc = 0 Then
        PythonLauncher = "python"
        Exit Function
    End If
    rc = shell.Run("cmd /d /c py -3 --version >nul 2>&1", 0, True)
    If rc = 0 Then
        PythonLauncher = "py -3"
        Exit Function
    End If
    Err.Raise vbObjectError + 2811, "PythonLauncher", "Δεν βρέθηκε διαθέσιμο Python runtime."
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

Private Function RunBridge(ByVal arguments As String) As String
    Dim shell As Object
    Dim tempFile As String
    Dim command As String
    Dim rc As Long

    tempFile = Environ$("TEMP") & "\eka_patient_scheduling_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    command = "cmd /d /c set PYTHONIOENCODING=utf-8&& " & PythonLauncher() & " " & _
              QuoteArg(BridgePath()) & " " & arguments & " > " & QuoteArg(tempFile) & " 2>&1"

    Set shell = CreateObject("WScript.Shell")
    rc = shell.Run(command, 0, True)
    RunBridge = ReadUtf8Text(tempFile)
    On Error Resume Next
    Kill tempFile
    On Error GoTo 0

    If rc <> 0 And Left$(Trim$(RunBridge), 5) <> "ERROR" Then
        Err.Raise vbObjectError + 2812, "RunBridge", Trim$(RunBridge)
    End If
End Function

Public Function LoadRecurringSchedulingOptions( _
    ByVal patientId As String, _
    ByVal treatment As String, _
    ByVal dayPattern As String _
) As String
    Dim args As String
    args = "recurring-vba " & QuoteArg(ThisWorkbook.FullName) & _
           " --patient-id " & QuoteArg(patientId) & _
           " --date " & QuoteArg(Format$(Date, "yyyy-mm-dd")) & _
           " --treatment " & QuoteArg(treatment) & _
           " --day-pattern " & QuoteArg(dayPattern)
    LoadRecurringSchedulingOptions = RunBridge(args)
End Function

Public Function ProposeAutomaticSchedule( _
    ByVal patientId As String, _
    ByVal treatment As String, _
    ByVal dayPattern As String _
) As String
    Dim args As String
    args = "auto-propose-vba " & QuoteArg(ThisWorkbook.FullName) & _
           " --patient-id " & QuoteArg(patientId) & _
           " --date " & QuoteArg(Format$(Date, "yyyy-mm-dd")) & _
           " --request " & QuoteArg(Trim$(treatment) & "|" & Trim$(dayPattern))
    ProposeAutomaticSchedule = RunBridge(args)
End Function

Public Function ValidateRecurringSelection( _
    ByVal patientId As String, _
    ByVal treatment As String, _
    ByVal dayPattern As String, _
    ByVal therapistId As String, _
    ByVal startTime As String _
) As String
    Dim args As String
    args = "validate-recurring-vba " & QuoteArg(ThisWorkbook.FullName) & _
           " --patient-id " & QuoteArg(patientId) & _
           " --date " & QuoteArg(Format$(Date, "yyyy-mm-dd")) & _
           " --treatment " & QuoteArg(treatment) & _
           " --day-pattern " & QuoteArg(dayPattern) & _
           " --therapist-id " & QuoteArg(therapistId) & _
           " --start-time " & QuoteArg(startTime)
    ValidateRecurringSelection = RunBridge(args)
End Function

Public Sub SavePatientRecurringAssignment( _
    ByVal patientId As String, _
    ByVal treatment As String, _
    ByVal dayPattern As String, _
    ByVal therapistId As String, _
    ByVal startTime As String _
)
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim timeCol As Long
    Dim daysCol As Long
    Dim therapistCol As Long
    Dim prefix As String
    Dim previousEvents As Boolean
    Dim errorText As String

    previousEvents = Application.EnableEvents
    On Error GoTo SaveAssignmentError

    Set ws = ThisWorkbook.Worksheets("PATIENT_PLANNER")
    rowIndex = FindPlannerPatientRow(ws, patientId)
    If rowIndex = 0 Then Err.Raise vbObjectError + 2813, , "Δεν βρέθηκε ο ασθενής στο PATIENT_PLANNER."

    Select Case Trim$(treatment)
        Case "ΦΘ": prefix = "ΦΘ"
        Case "Ρομποτικό": prefix = "Ρομποτικό"
        Case Else
            Err.Raise vbObjectError + 2814, , "Η θεραπεία δεν υποστηρίζεται ακόμη από το physiotherapist scheduling flow."
    End Select

    timeCol = FindPlannerHeader(ws, prefix & "_Ώρα")
    daysCol = FindPlannerHeader(ws, prefix & "_Ημέρες")
    therapistCol = FindPlannerHeader(ws, prefix & "_Θεραπευτής")

    ws.Cells(rowIndex, timeCol).Value = TimeValue(startTime)
    ws.Cells(rowIndex, timeCol).NumberFormat = "hh:mm"
    ws.Cells(rowIndex, daysCol).Value = dayPattern
    ws.Cells(rowIndex, therapistCol).Value = therapistId
    ws.Rows(rowIndex).Calculate

    Application.EnableEvents = False
    ThisWorkbook.Save
    Application.EnableEvents = previousEvents
    Exit Sub

SaveAssignmentError:
    errorText = Err.Description
    On Error Resume Next
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 2817, "SavePatientRecurringAssignment", errorText
End Sub

Private Function FindPlannerPatientRow(ByVal ws As Worksheet, ByVal patientId As String) As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 1).Value)), Trim$(patientId), vbTextCompare) = 0 Then
            FindPlannerPatientRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FindPlannerHeader(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim colIndex As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For colIndex = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, colIndex).Value)), headerText, vbTextCompare) = 0 Then
            FindPlannerHeader = colIndex
            Exit Function
        End If
    Next colIndex
    Err.Raise vbObjectError + 2815, , "Λείπει στήλη PATIENT_PLANNER: " & headerText
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
    FindPatientDisplayName = ""
End Function

Public Sub OpenPatientScheduling(ByVal patientId As String)
    On Error GoTo OpenError
    Load frmPatientScheduling
    If frmPatientScheduling.Visible Then Exit Sub
    frmPatientScheduling.Configure patientId
    frmPatientScheduling.Show vbModeless
    Exit Sub
OpenError:
    Err.Raise vbObjectError + 2816, "OpenPatientScheduling", Err.Description
End Sub
'''

PATIENT_SCHEDULING_FORM_CODE = r'''Option Explicit

Private mPatientId As String
Private mSlotsByTherapist As Object

Public Sub Configure(ByVal patientId As String)
    Dim patientName As String

    mPatientId = Trim$(patientId)
    patientName = FindPatientDisplayName(mPatientId)
    If Len(patientName) > 0 Then
        lblPatient.Caption = "Ασθενής: " & patientName & "   |   Patient ID: " & mPatientId
    Else
        lblPatient.Caption = "Ασθενής: Patient ID " & mPatientId
    End If
    LoadSelections
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Προγραμματισμός θεραπείας"
    Me.Width = 560
    Me.Height = 500
    Me.BackColor = RGB(245, 247, 250)

    lblTitle.Caption = "Πρόγραμμα θεραπείας ασθενούς"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 15
    lblInfo.Caption = "Επίλεξε θεραπεία και ημέρες για να εμφανιστούν οι διαθέσιμες επιλογές."

    cboTreatment.Style = 2
    cboDayPattern.Style = 2
    cboTherapist.Style = 2
    cboTherapist.ColumnCount = 2
    cboTherapist.ColumnWidths = "230 pt;0 pt"
    cboTime.Style = 2

    Set mSlotsByTherapist = CreateObject("Scripting.Dictionary")
    LoadSettingsChoices
End Sub

Private Sub LoadSettingsChoices()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim valueText As String

    Set ws = ThisWorkbook.Worksheets("SETTINGS")
    cboTreatment.Clear
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    For rowIndex = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(rowIndex, 5).Value))
        If valueText = "ΦΘ" Or valueText = "Ρομποτικό" Then cboTreatment.AddItem valueText
    Next rowIndex

    cboDayPattern.Clear
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    For rowIndex = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(rowIndex, 4).Value))
        If Len(valueText) > 0 Then cboDayPattern.AddItem valueText
    Next rowIndex
End Sub

Private Sub cboTreatment_Change()
    lblProposal.Caption = ""
    LoadSelections
End Sub

Private Sub cboDayPattern_Change()
    lblProposal.Caption = ""
    LoadSelections
End Sub

Private Sub cmdRefresh_Click()
    LoadSelections
End Sub

Private Sub LoadSelections()
    Dim response As String
    Dim lines() As String
    Dim fields() As String
    Dim lineIndex As Long

    cboTherapist.Clear
    cboTime.Clear
    Set mSlotsByTherapist = CreateObject("Scripting.Dictionary")

    If Len(mPatientId) = 0 Then Exit Sub
    If Len(Trim$(cboTreatment.Value)) = 0 Then Exit Sub
    If Len(Trim$(cboDayPattern.Value)) = 0 Then Exit Sub

    On Error GoTo LoadError
    response = LoadRecurringSchedulingOptions( _
        mPatientId, _
        Trim$(cboTreatment.Value), _
        Trim$(cboDayPattern.Value) _
    )
    lines = Split(Replace(response, vbCrLf, vbLf), vbLf)
    For lineIndex = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(lineIndex))) = 0 Then GoTo NextLine
        fields = Split(lines(lineIndex), vbTab)
        If fields(0) = "ERROR" Then Err.Raise vbObjectError + 2820, , fields(1)
        If fields(0) = "THERAPIST" And UBound(fields) >= 4 Then
            cboTherapist.AddItem fields(3)
            cboTherapist.List(cboTherapist.ListCount - 1, 1) = fields(2)
            mSlotsByTherapist(fields(2)) = fields(4)
        End If
NextLine:
    Next lineIndex

    If cboTherapist.ListCount = 0 Then
        lblInfo.Caption = "Δεν υπάρχει κοινή διαθέσιμη ώρα για όλες τις επιλεγμένες ημέρες."
    Else
        lblInfo.Caption = "Οι θεραπευτές εμφανίζονται με βάση τη διαθεσιμότητα και τον τρέχοντα φόρτο."
    End If
    Exit Sub
LoadError:
    lblInfo.Caption = "Δεν ήταν δυνατή η φόρτωση της διαθεσιμότητας. Δοκίμασε Ανανέωση."
End Sub

Private Sub cboTherapist_Change()
    Dim therapistId As String
    Dim slots() As String
    Dim index As Long

    cboTime.Clear
    If cboTherapist.ListIndex < 0 Then Exit Sub
    therapistId = CStr(cboTherapist.List(cboTherapist.ListIndex, 1))
    If Not mSlotsByTherapist.Exists(therapistId) Then Exit Sub
    slots = Split(CStr(mSlotsByTherapist(therapistId)), "|")
    For index = LBound(slots) To UBound(slots)
        If Len(Trim$(slots(index))) > 0 Then cboTime.AddItem Trim$(slots(index))
    Next index
End Sub

Private Sub cmdAutoPropose_Click()
    Dim response As String
    Dim lines() As String
    Dim fields() As String
    Dim lineIndex As Long
    Dim therapistIndex As Long
    Dim proposedTherapistId As String
    Dim proposedTherapistName As String
    Dim proposedTime As String

    If Len(Trim$(cboTreatment.Value)) = 0 Or Len(Trim$(cboDayPattern.Value)) = 0 Then
        MsgBox "Επίλεξε θεραπεία και ημέρες πριν ζητήσεις πρόταση.", vbExclamation, "Αυτόματη πρόταση"
        Exit Sub
    End If

    On Error GoTo ProposalError
    response = ProposeAutomaticSchedule( _
        mPatientId, _
        Trim$(cboTreatment.Value), _
        Trim$(cboDayPattern.Value) _
    )
    lines = Split(Replace(response, vbCrLf, vbLf), vbLf)
    For lineIndex = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(lineIndex))) = 0 Then GoTo NextProposalLine
        fields = Split(lines(lineIndex), vbTab)
        If fields(0) = "ERROR" Then Err.Raise vbObjectError + 2822, , fields(1)
        If fields(0) = "ASSIGNMENT" And UBound(fields) >= 6 Then
            proposedTherapistId = fields(3)
            proposedTherapistName = fields(4)
            proposedTime = fields(5)
            Exit For
        End If
NextProposalLine:
    Next lineIndex

    If Len(proposedTherapistId) = 0 Then
        lblProposal.Caption = "Δεν βρέθηκε ασφαλής αυτόματη πρόταση για αυτή την επιλογή."
        Exit Sub
    End If

    LoadSelections
    For therapistIndex = 0 To cboTherapist.ListCount - 1
        If StrComp(CStr(cboTherapist.List(therapistIndex, 1)), proposedTherapistId, vbTextCompare) = 0 Then
            cboTherapist.ListIndex = therapistIndex
            Exit For
        End If
    Next therapistIndex
    cboTime.Value = proposedTime
    lblProposal.Caption = "Πρόταση: " & proposedTherapistName & " στις " & proposedTime & ". Έλεγξέ την και πάτησε Αποθήκευση για επιβεβαίωση."
    Exit Sub
ProposalError:
    lblProposal.Caption = "Δεν ήταν δυνατή η δημιουργία αυτόματης πρότασης."
End Sub

Private Sub cmdSave_Click()
    Dim therapistId As String
    Dim response As String
    Dim fields() As String

    If cboTherapist.ListIndex < 0 Or Len(Trim$(cboTime.Value)) = 0 Then
        MsgBox "Επίλεξε θεραπευτή και ώρα.", vbExclamation, "Προγραμματισμός"
        Exit Sub
    End If

    therapistId = CStr(cboTherapist.List(cboTherapist.ListIndex, 1))
    On Error GoTo SaveError
    response = ValidateRecurringSelection( _
        mPatientId, _
        Trim$(cboTreatment.Value), _
        Trim$(cboDayPattern.Value), _
        therapistId, _
        Trim$(cboTime.Value) _
    )
    fields = Split(Trim$(response), vbTab)
    If fields(0) = "ERROR" Then Err.Raise vbObjectError + 2821, , fields(1)
    If UBound(fields) < 2 Or fields(1) <> "1" Then
        MsgBox IIf(UBound(fields) >= 2, fields(2), "Η επιλογή δεν είναι πλέον διαθέσιμη."), vbExclamation, "Επανέλεγχος"
        LoadSelections
        Exit Sub
    End If

    SavePatientRecurringAssignment _
        mPatientId, _
        Trim$(cboTreatment.Value), _
        Trim$(cboDayPattern.Value), _
        therapistId, _
        Trim$(cboTime.Value)

    MsgBox "Η θεραπεία προγραμματίστηκε επιτυχώς.", vbInformation, "Προγραμματισμός"
    lblProposal.Caption = "Η θεραπεία αποθηκεύτηκε. Μπορείς να επιλέξεις την επόμενη θεραπεία ή Κλείσιμο."
    cboTreatment.SetFocus
    Exit Sub
SaveError:
    MsgBox Err.Description, vbCritical, "Προγραμματισμός"
End Sub

Private Sub cmdCancel_Click()
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


def _ensure_control(designer, prog_id: str, name: str, caption: str | None, top: float, *, position_control):
    try:
        control = designer.Controls(name)
    except Exception:
        control = designer.Controls.Add(prog_id, name, True)
    if caption is not None:
        control.Caption = caption

    left = 24
    width = 240
    height = 18
    if name == "lblTitle":
        width = 430
        height = 26
    elif name == "lblPatient":
        width = 430
        height = 22
    elif name in {"lblInfo", "lblProposal"}:
        width = 430
        height = 30
        try:
            control.WordWrap = True
        except Exception:
            pass
    elif name.startswith("lbl"):
        width = 300
        height = 18
    elif name.startswith("cbo"):
        width = 260
        height = 20
    elif name in {"cmdAutoPropose", "cmdRefresh"}:
        width = 220
        height = 28
    elif name == "cmdCancel":
        width = 120
        height = 30
    elif name == "cmdSave":
        left = 160
        width = 120
        height = 30

    position_control(control, left, top)
    try:
        control.Width = width
    except Exception:
        pass
    try:
        control.Height = height
    except Exception:
        pass
    return control


def install_patient_scheduling_ui(vbproject, *, position_control) -> None:
    """Install or update the scheduling runtime/form without deleting mature components."""

    module = _get_component(vbproject, PATIENT_SCHEDULING_MODULE_NAME)
    if module is None:
        module = vbproject.VBComponents.Add(1)
        module.Name = PATIENT_SCHEDULING_MODULE_NAME
    _replace_code(module, PATIENT_SCHEDULING_MODULE_CODE)

    form = _get_component(vbproject, PATIENT_SCHEDULING_FORM_NAME)
    if form is None:
        form = vbproject.VBComponents.Add(3)
        form.Name = PATIENT_SCHEDULING_FORM_NAME
    designer = form.Designer
    designer.Caption = "Προγραμματισμός θεραπείας"
    try:
        designer.Width = 520
        designer.Height = 470
    except Exception:
        pass

    controls = (
        ("Forms.Label.1", "lblTitle", "Πρόγραμμα θεραπείας ασθενούς", 18),
        ("Forms.Label.1", "lblPatient", "Ασθενής:", 52),
        ("Forms.Label.1", "lblTreatment", "1. Θεραπεία", 86),
        ("Forms.ComboBox.1", "cboTreatment", None, 102),
        ("Forms.Label.1", "lblDayPattern", "2. Ημέρες θεραπείας", 136),
        ("Forms.ComboBox.1", "cboDayPattern", None, 152),
        ("Forms.Label.1", "lblTherapist", "3. Διαθέσιμος θεραπευτής", 186),
        ("Forms.ComboBox.1", "cboTherapist", None, 202),
        ("Forms.Label.1", "lblTime", "4. Διαθέσιμη ώρα", 236),
        ("Forms.ComboBox.1", "cboTime", None, 252),
        ("Forms.Label.1", "lblInfo", "Επίλεξε θεραπεία και ημέρες για να εμφανιστούν οι διαθέσιμες επιλογές.", 286),
        ("Forms.Label.1", "lblProposal", "", 316),
        ("Forms.CommandButton.1", "cmdAutoPropose", "Πρότεινε πρόγραμμα", 348),
        ("Forms.CommandButton.1", "cmdRefresh", "Ανανέωση διαθεσιμότητας", 382),
        ("Forms.CommandButton.1", "cmdCancel", "Κλείσιμο", 424),
        ("Forms.CommandButton.1", "cmdSave", "Αποθήκευση", 424),
    )
    for prog_id, name, caption, top in controls:
        _ensure_control(
            designer,
            prog_id,
            name,
            caption,
            top,
            position_control=position_control,
        )

    _replace_code(form, PATIENT_SCHEDULING_FORM_CODE)
