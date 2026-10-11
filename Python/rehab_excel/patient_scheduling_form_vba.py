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
    ThisWorkbook.Save
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

Public Sub OpenPatientScheduling(ByVal patientId As String)
    Load frmPatientScheduling
    frmPatientScheduling.Configure patientId
    frmPatientScheduling.Show
End Sub
'''

PATIENT_SCHEDULING_FORM_CODE = r'''Option Explicit

Private mPatientId As String
Private mSlotsByTherapist As Object

Public Sub Configure(ByVal patientId As String)
    mPatientId = Trim$(patientId)
    lblPatient.Caption = "Patient ID: " & mPatientId
    LoadSelections
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Προγραμματισμός θεραπείας"
    Me.Width = 520
    Me.Height = 430
    Me.BackColor = RGB(245, 247, 250)

    lblTitle.Caption = "Θεραπείες και πρόγραμμα"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 15

    cboTreatment.Style = 2
    cboDayPattern.Style = 2
    cboTherapist.Style = 2
    cboTherapist.ColumnCount = 2
    cboTherapist.ColumnWidths = "190 pt;0 pt"
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
    LoadSelections
End Sub

Private Sub cboDayPattern_Change()
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
    response = LoadRecurringSchedulingOptions(
        mPatientId,
        Trim$(cboTreatment.Value),
        Trim$(cboDayPattern.Value)
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
        lblInfo.Caption = "Δεν υπάρχει κοινή διαθέσιμη ώρα για όλες τις ημέρες του μοτίβου."
    Else
        lblInfo.Caption = "Οι θεραπευτές είναι ταξινομημένοι με βάση τον τρέχοντα φόρτο."
    End If
    Exit Sub
LoadError:
    lblInfo.Caption = Err.Description
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
    response = ValidateRecurringSelection(
        mPatientId,
        Trim$(cboTreatment.Value),
        Trim$(cboDayPattern.Value),
        therapistId,
        Trim$(cboTime.Value)
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
    Unload Me
    Exit Sub
SaveError:
    MsgBox Err.Description, vbCritical, "Προγραμματισμός"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'''


def install_patient_scheduling_ui(vbproject, *, position_control) -> None:
    for component_name in (PATIENT_SCHEDULING_FORM_NAME, PATIENT_SCHEDULING_MODULE_NAME):
        try:
            existing = vbproject.VBComponents(component_name)
        except Exception:
            existing = None
        if existing is not None:
            vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = PATIENT_SCHEDULING_MODULE_NAME
    module.CodeModule.AddFromString(PATIENT_SCHEDULING_MODULE_CODE)

    form = vbproject.VBComponents.Add(3)
    form.Name = PATIENT_SCHEDULING_FORM_NAME
    designer = form.Designer
    designer.Caption = "Προγραμματισμός θεραπείας"

    controls = (
        ("Forms.Label.1", "lblTitle", "Θεραπείες και πρόγραμμα", 18),
        ("Forms.Label.1", "lblPatient", "Patient ID:", 52),
        ("Forms.Label.1", "lblTreatment", "Θεραπεία", 86),
        ("Forms.ComboBox.1", "cboTreatment", None, 82),
        ("Forms.Label.1", "lblDayPattern", "Ημέρες", 126),
        ("Forms.ComboBox.1", "cboDayPattern", None, 122),
        ("Forms.Label.1", "lblTherapist", "Θεραπευτής", 166),
        ("Forms.ComboBox.1", "cboTherapist", None, 162),
        ("Forms.Label.1", "lblTime", "Ώρα", 206),
        ("Forms.ComboBox.1", "cboTime", None, 202),
        ("Forms.Label.1", "lblInfo", "", 244),
        ("Forms.CommandButton.1", "cmdRefresh", "Ανανέωση διαθεσιμότητας", 300),
        ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 342),
        ("Forms.CommandButton.1", "cmdSave", "Αποθήκευση προγράμματος", 342),
    )
    for prog_id, name, caption, top in controls:
        control = designer.Controls.Add(prog_id, name, True)
        if caption is not None:
            control.Caption = caption
        position_control(control, 24, top)

    form.CodeModule.AddFromString(PATIENT_SCHEDULING_FORM_CODE)
