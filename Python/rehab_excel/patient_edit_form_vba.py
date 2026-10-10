from __future__ import annotations

PATIENT_EDIT_FORM_NAME = "frmEditPatient"

PATIENT_EDIT_FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Επεξεργασία ασθενή"
        .Width = 520
        .Height = 650
        .StartUpPosition = 1
        .BackColor = RGB(245, 247, 250)
    End With

    StyleTitle lblTitle, "Επεξεργασία ασθενή", 18

    StyleLabel lblPatient, "Ασθενής *", 62
    StyleComboBox cboPatient, 58

    StyleLabel lblPatientID, "Patient ID", 106
    StyleTextBox txtPatientID, 102
    txtPatientID.Locked = True
    txtPatientID.TabStop = False
    txtPatientID.BackColor = RGB(238, 242, 247)

    StyleLabel lblHospitalMRN, "ΑΜ Νοσοκομείου", 150
    StyleTextBox txtHospitalMRN, 146

    StyleLabel lblDisplayName, "Ονοματεπώνυμο *", 194
    StyleTextBox txtDisplayName, 190

    StyleLabel lblResponsibleDoctor, "Υπεύθυνος γιατρός", 238
    StyleTextBox txtResponsibleDoctor, 234

    StyleLabel lblRoom, "Θάλαμος", 282
    StyleComboBox cboRoom, 278

    StyleLabel lblInfectious, "Λοιμώδης", 326
    StyleCheckBox chkInfectious, "Ναι", 322

    StyleLabel lblStatus, "Κατάσταση *", 370
    StyleComboBox cboStatus, 366

    With lblInfo
        .Caption = "Οι μόνιμες αλλαγές στοιχείων ενημερώνουν το μητρώο και τις προβολές. Το Patient ID δεν αλλάζει."
        .Left = 38
        .Top = 414
        .Width = 430
        .Height = 42
        .WordWrap = True
        .Font.Name = "Calibri"
        .Font.Size = 9
        .ForeColor = RGB(71, 85, 105)
        .BackStyle = 0
    End With

    With cmdSave
        .Caption = "Αποθήκευση αλλαγών"
        .Left = 255
        .Top = 505
        .Width = 175
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Default = True
    End With

    With cmdCancel
        .Caption = "Ακύρωση"
        .Left = 95
        .Top = 505
        .Width = 130
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Cancel = True
    End With

    LoadSettingsValues
    LoadPatients
End Sub

Private Sub StyleTitle(ByVal control As MSForms.Label, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 38
        .Top = topPosition
        .Width = 430
        .Height = 26
        .TextAlign = 2
        .Font.Name = "Calibri"
        .Font.Size = 15
        .Font.Bold = True
        .ForeColor = RGB(45, 55, 72)
        .BackStyle = 0
    End With
End Sub

Private Sub StyleLabel(ByVal control As MSForms.Label, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 38
        .Top = topPosition
        .Width = 145
        .Height = 20
        .Font.Name = "Calibri"
        .Font.Size = 10
        .ForeColor = RGB(51, 65, 85)
        .BackStyle = 0
    End With
End Sub

Private Sub StyleTextBox(ByVal control As MSForms.TextBox, ByVal topPosition As Single)
    With control
        .Left = 190
        .Top = topPosition
        .Width = 275
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub

Private Sub StyleComboBox(ByVal control As MSForms.ComboBox, ByVal topPosition As Single)
    With control
        .Left = 190
        .Top = topPosition
        .Width = 275
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Style = 2
    End With
End Sub

Private Sub StyleCheckBox(ByVal control As MSForms.CheckBox, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 190
        .Top = topPosition
        .Width = 80
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Value = False
    End With
End Sub

Private Sub LoadSettingsValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim valueText As String

    Set ws = ThisWorkbook.Worksheets("SETTINGS")

    cboRoom.Clear
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    For rowIndex = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(rowIndex, 8).Value))
        If Len(valueText) > 0 Then cboRoom.AddItem valueText
    Next rowIndex

    cboStatus.Clear
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    For rowIndex = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(rowIndex, 6).Value))
        If Len(valueText) > 0 Then cboStatus.AddItem valueText
    Next rowIndex
End Sub

Private Sub LoadPatients()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim displayName As String
    Dim patientId As String

    Set ws = ThisWorkbook.Worksheets("PATIENTS")
    cboPatient.Clear
    cboPatient.ColumnCount = 2
    cboPatient.ColumnWidths = "310 pt;0 pt"

    lastRow = Application.Max(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, ws.Cells(ws.Rows.Count, 3).End(xlUp).Row)
    For rowIndex = 2 To lastRow
        patientId = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
        displayName = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
        If Len(patientId) > 0 And Len(displayName) > 0 Then
            cboPatient.AddItem displayName & "  [" & patientId & "]"
            cboPatient.List(cboPatient.ListCount - 1, 1) = patientId
        End If
    Next rowIndex
End Sub

Private Sub cboPatient_Change()
    Dim patientId As String
    If cboPatient.ListIndex < 0 Then Exit Sub
    patientId = CStr(cboPatient.List(cboPatient.ListIndex, 1))
    LoadPatientById patientId
End Sub

Private Sub LoadPatientById(ByVal patientId As String)
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim typeCol As Long
    Dim mrnCol As Long
    Dim doctorCol As Long
    Dim patientType As String

    Set ws = ThisWorkbook.Worksheets("PATIENTS")
    rowIndex = FindPatientRow(ws, patientId)
    If rowIndex = 0 Then Exit Sub

    typeCol = FindHeader(ws, Array("PatientType", "ΤύποςΑσθενή", "Τύπος Ασθενή"))
    mrnCol = FindHeader(ws, Array("HospitalMRN", "ΑΜ Νοσοκομείου", "ΑΜΝοσοκομείου"))
    doctorCol = FindHeader(ws, Array("ResponsibleDoctor", "ΥπεύθυνοςΙατρός", "Υπεύθυνος Ιατρός", "ΥπεύθυνοςΓιατρός", "Υπεύθυνος Γιατρός"))

    txtPatientID.Text = CStr(ws.Cells(rowIndex, 1).Value)
    txtDisplayName.Text = CStr(ws.Cells(rowIndex, 3).Value)
    cboRoom.Value = CStr(ws.Cells(rowIndex, 2).Value)
    chkInfectious.Value = IsTruthy(ws.Cells(rowIndex, 4).Value)
    cboStatus.Value = CStr(ws.Cells(rowIndex, 5).Value)

    If mrnCol > 0 Then txtHospitalMRN.Text = CStr(ws.Cells(rowIndex, mrnCol).Value)
    If doctorCol > 0 Then txtResponsibleDoctor.Text = CStr(ws.Cells(rowIndex, doctorCol).Value)

    If typeCol > 0 Then patientType = Trim$(CStr(ws.Cells(rowIndex, typeCol).Value))
    ApplyPatientTypeRules patientType
End Sub

Private Sub ApplyPatientTypeRules(ByVal patientType As String)
    Dim isOutpatient As Boolean
    isOutpatient = (LCase$(Trim$(patientType)) = LCase$("Εξωτερικός"))

    lblRoom.Enabled = Not isOutpatient
    cboRoom.Enabled = Not isOutpatient
    lblInfectious.Enabled = Not isOutpatient
    chkInfectious.Enabled = Not isOutpatient
    lblResponsibleDoctor.Enabled = Not isOutpatient
    txtResponsibleDoctor.Enabled = Not isOutpatient

    If isOutpatient Then
        cboRoom.ListIndex = -1
        chkInfectious.Value = False
        txtResponsibleDoctor.Text = ""
    End If
End Sub

Private Function FindPatientRow(ByVal ws As Worksheet, ByVal patientId As String) As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 1).Value)), Trim$(patientId), vbTextCompare) = 0 Then
            FindPatientRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FindHeader(ByVal ws As Worksheet, ByVal candidates As Variant) As Long
    Dim lastCol As Long
    Dim col As Long
    Dim candidate As Variant
    Dim text As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        text = Trim$(CStr(ws.Cells(1, col).Value))
        For Each candidate In candidates
            If StrComp(text, CStr(candidate), vbTextCompare) = 0 Then
                FindHeader = col
                Exit Function
            End If
        Next candidate
    Next col
End Function

Private Function IsTruthy(ByVal value As Variant) As Boolean
    Dim text As String
    If VarType(value) = vbBoolean Then
        IsTruthy = CBool(value)
        Exit Function
    End If
    If IsNumeric(value) Then
        IsTruthy = (CDbl(value) <> 0)
        Exit Function
    End If
    text = LCase$(Trim$(CStr(value)))
    IsTruthy = (text = "ν" Or text = "ναι" Or text = "yes" Or text = "true" Or text = "1")
End Function

Private Function ValidateForm() As Boolean
    If cboPatient.ListIndex < 0 Then
        MsgBox "Επιλέξτε ασθενή.", vbExclamation, "Επεξεργασία ασθενή"
        Exit Function
    End If
    If Len(Trim$(txtDisplayName.Text)) = 0 Then
        MsgBox "Το ονοματεπώνυμο είναι υποχρεωτικό.", vbExclamation, "Επεξεργασία ασθενή"
        txtDisplayName.SetFocus
        Exit Function
    End If
    If Len(Trim$(cboStatus.Value)) = 0 Then
        MsgBox "Η κατάσταση είναι υποχρεωτική.", vbExclamation, "Επεξεργασία ασθενή"
        cboStatus.SetFocus
        Exit Function
    End If
    If cboRoom.Enabled And Len(Trim$(cboRoom.Value)) = 0 Then
        MsgBox "Ο θάλαμος είναι υποχρεωτικός για εσωτερικό ασθενή.", vbExclamation, "Επεξεργασία ασθενή"
        cboRoom.SetFocus
        Exit Function
    End If
    If txtResponsibleDoctor.Enabled And Len(Trim$(txtResponsibleDoctor.Text)) = 0 Then
        MsgBox "Ο υπεύθυνος γιατρός είναι υποχρεωτικός για εσωτερικό ασθενή.", vbExclamation, "Επεξεργασία ασθενή"
        txtResponsibleDoctor.SetFocus
        Exit Function
    End If
    ValidateForm = True
End Function

Private Sub cmdSave_Click()
    If Not ValidateForm() Then Exit Sub

    On Error GoTo SaveError
    UpdatePatientInWorkbook _
        Trim$(txtPatientID.Text), _
        Trim$(txtHospitalMRN.Text), _
        Trim$(txtDisplayName.Text), _
        Trim$(cboRoom.Value), _
        CBool(chkInfectious.Value), _
        Trim$(txtResponsibleDoctor.Text), _
        Trim$(cboStatus.Value)

    MsgBox "Οι αλλαγές αποθηκεύτηκαν επιτυχώς.", vbInformation, "Επεξεργασία ασθενή"
    Unload Me
    On Error Resume Next
    Unload frmRegistrationMenu
    On Error GoTo 0
    GoToMaster
    Exit Sub

SaveError:
    MsgBox "Η ενημέρωση δεν ολοκληρώθηκε: " & Err.Description, vbCritical, "Επεξεργασία ασθενή"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'''


def install_patient_edit_form(vbproject, *, position_control) -> None:
    try:
        existing = vbproject.VBComponents(PATIENT_EDIT_FORM_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    form = vbproject.VBComponents.Add(3)
    form.Name = PATIENT_EDIT_FORM_NAME
    designer = form.Designer
    designer.Caption = "Επεξεργασία ασθενή"

    controls = (
        ("Forms.Label.1", "lblTitle", "Επεξεργασία ασθενή", 18),
        ("Forms.Label.1", "lblPatient", "Ασθενής *", 62),
        ("Forms.ComboBox.1", "cboPatient", None, 58),
        ("Forms.Label.1", "lblPatientID", "Patient ID", 106),
        ("Forms.TextBox.1", "txtPatientID", None, 102),
        ("Forms.Label.1", "lblHospitalMRN", "ΑΜ Νοσοκομείου", 150),
        ("Forms.TextBox.1", "txtHospitalMRN", None, 146),
        ("Forms.Label.1", "lblDisplayName", "Ονοματεπώνυμο *", 194),
        ("Forms.TextBox.1", "txtDisplayName", None, 190),
        ("Forms.Label.1", "lblResponsibleDoctor", "Υπεύθυνος γιατρός", 238),
        ("Forms.TextBox.1", "txtResponsibleDoctor", None, 234),
        ("Forms.Label.1", "lblRoom", "Θάλαμος", 282),
        ("Forms.ComboBox.1", "cboRoom", None, 278),
        ("Forms.Label.1", "lblInfectious", "Λοιμώδης", 326),
        ("Forms.CheckBox.1", "chkInfectious", "Ναι", 322),
        ("Forms.Label.1", "lblStatus", "Κατάσταση", 370),
        ("Forms.ComboBox.1", "cboStatus", None, 366),
        ("Forms.Label.1", "lblInfo", "", 414),
        ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 505),
        ("Forms.CommandButton.1", "cmdSave", "Αποθήκευση αλλαγών", 505),
    )

    for prog_id, name, caption, top in controls:
        control = designer.Controls.Add(prog_id, name, True)
        if caption is not None:
            control.Caption = caption
        position_control(control, 24, top)

    form.CodeModule.AddFromString(PATIENT_EDIT_FORM_CODE)
