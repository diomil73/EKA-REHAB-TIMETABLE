from __future__ import annotations

PATIENT_FORM_NAME = "frmNewPatient"

PATIENT_FORM_CODE = '''Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Νέος ασθενής"
        .Width = 500
        .Height = 560
        .StartUpPosition = 1
        .BackColor = RGB(245, 247, 250)
    End With

    StyleTitle lblTitle, "Εγγραφή νέου ασθενή", 18

    StyleLabel lblPatientType, "Τύπος ασθενή *", 62
    StyleComboBox cboPatientType, 58
    cboPatientType.AddItem "Εσωτερικός"
    cboPatientType.AddItem "Εξωτερικός"
    cboPatientType.ListIndex = 0

    StyleLabel lblPatientID, "Patient ID", 106
    StyleTextBox txtPatientID, 102
    txtPatientID.Text = "Αυτόματο κατά την αποθήκευση"
    txtPatientID.Locked = True
    txtPatientID.TabStop = False
    txtPatientID.BackColor = RGB(238, 242, 247)

    StyleLabel lblHospitalMRN, "ΑΜ Νοσοκομείου", 150
    StyleTextBox txtHospitalMRN, 146

    StyleLabel lblDisplayName, "Ονοματεπώνυμο *", 194
    StyleTextBox txtDisplayName, 190

    StyleLabel lblRoom, "Θάλαμος", 238
    StyleComboBox cboRoom, 234

    StyleLabel lblInfectious, "Λοιμώδης", 282
    StyleCheckBox chkInfectious, "Ναι", 278

    StyleLabel lblStatus, "Κατάσταση", 326
    StyleComboBox cboStatus, 322

    With lblInfo
        .Caption = "Το Patient ID είναι μόνιμος εσωτερικός κωδικός. Ο ΑΜ Νοσοκομείου μπορεί να συμπληρωθεί και αργότερα."
        .Left = 38
        .Top = 370
        .Width = 420
        .Height = 38
        .WordWrap = True
        .Font.Name = "Calibri"
        .Font.Size = 9
        .ForeColor = RGB(71, 85, 105)
        .BackStyle = 0
    End With

    With lblRequired
        .Caption = "* Υποχρεωτικό πεδίο"
        .Left = 38
        .Top = 414
        .Width = 180
        .Height = 18
        .Font.Name = "Calibri"
        .Font.Size = 9
        .ForeColor = RGB(100, 116, 139)
        .BackStyle = 0
    End With

    With cmdSave
        .Caption = "Έλεγχος στοιχείων"
        .Left = 235
        .Top = 452
        .Width = 170
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Default = True
    End With

    With cmdCancel
        .Caption = "Ακύρωση"
        .Left = 92
        .Top = 452
        .Width = 125
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Cancel = True
    End With

    LoadSettingsValues
    ApplyPatientTypeRules
    cboPatientType.SetFocus
End Sub

Private Sub StyleTitle(ByVal control As MSForms.Label, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 38
        .Top = topPosition
        .Width = 420
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
        .Width = 265
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub

Private Sub StyleComboBox(ByVal control As MSForms.ComboBox, ByVal topPosition As Single)
    With control
        .Left = 190
        .Top = topPosition
        .Width = 265
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

    On Error GoTo SettingsError
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
    Exit Sub

SettingsError:
    MsgBox "Δεν ήταν δυνατή η φόρτωση των επιλογών από το φύλλο SETTINGS.", vbExclamation, "Νέος ασθενής"
End Sub

Private Sub cboPatientType_Change()
    ApplyPatientTypeRules
End Sub

Private Sub ApplyPatientTypeRules()
    Dim isInpatient As Boolean
    isInpatient = (cboPatientType.Value <> "Εξωτερικός")

    lblRoom.Enabled = isInpatient
    cboRoom.Enabled = isInpatient
    lblInfectious.Enabled = isInpatient
    chkInfectious.Enabled = isInpatient
    lblStatus.Enabled = isInpatient
    cboStatus.Enabled = isInpatient

    If Not isInpatient Then
        cboRoom.ListIndex = -1
        chkInfectious.Value = False
        cboStatus.ListIndex = -1
    End If
End Sub

Private Function ValidateForm() As Boolean
    If Len(Trim$(cboPatientType.Value)) = 0 Then
        MsgBox "Ο τύπος ασθενή είναι υποχρεωτικός.", vbExclamation, "Έλεγχος στοιχείων"
        cboPatientType.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtDisplayName.Text)) = 0 Then
        MsgBox "Το ονοματεπώνυμο είναι υποχρεωτικό.", vbExclamation, "Έλεγχος στοιχείων"
        txtDisplayName.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

Private Sub cmdSave_Click()
    Dim message As String
    Dim patientType As String

    If Not ValidateForm() Then Exit Sub

    patientType = Trim$(cboPatientType.Value)
    message = "Τύπος: " & patientType & vbCrLf & _
              "Patient ID: Αυτόματο κατά την αποθήκευση" & vbCrLf & _
              "ΑΜ Νοσοκομείου: " & IIf(Len(Trim$(txtHospitalMRN.Text)) > 0, Trim$(txtHospitalMRN.Text), "-") & vbCrLf & _
              "Ονοματεπώνυμο: " & Trim$(txtDisplayName.Text) & vbCrLf

    If patientType = "Εσωτερικός" Then
        message = message & _
                  "Θάλαμος: " & IIf(Len(cboRoom.Value) > 0, cboRoom.Value, "-") & vbCrLf & _
                  "Λοιμώδης: " & IIf(chkInfectious.Value, "Ναι", "Όχι") & vbCrLf & _
                  "Κατάσταση: " & IIf(Len(cboStatus.Value) > 0, cboStatus.Value, "-") & vbCrLf
    Else
        message = message & "Εξωτερικός ασθενής: δεν εμφανίζεται στα φύλλα νοσηλευομένων." & vbCrLf
    End If

    message = message & vbCrLf & _
              "Τα στοιχεία είναι έτοιμα για αποστολή στο ασφαλές registration backend."

    MsgBox message, vbInformation, "Έλεγχος νέου ασθενή"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'''


def install_patient_form(vbproject, *, position_control) -> None:
    """Create the patient registration UserForm in an existing VBProject.

    Geometry is intentionally styled from VBA at runtime because the target
    Office build rejects COM Width/Height assignments.
    """

    try:
        existing = vbproject.VBComponents(PATIENT_FORM_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    form = vbproject.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = PATIENT_FORM_NAME
    designer = form.Designer
    designer.Caption = "Νέος ασθενής"

    controls = (
        ("Forms.Label.1", "lblTitle", "Εγγραφή νέου ασθενή", 18),
        ("Forms.Label.1", "lblPatientType", "Τύπος ασθενή *", 62),
        ("Forms.ComboBox.1", "cboPatientType", None, 58),
        ("Forms.Label.1", "lblPatientID", "Patient ID", 106),
        ("Forms.TextBox.1", "txtPatientID", None, 102),
        ("Forms.Label.1", "lblHospitalMRN", "ΑΜ Νοσοκομείου", 150),
        ("Forms.TextBox.1", "txtHospitalMRN", None, 146),
        ("Forms.Label.1", "lblDisplayName", "Ονοματεπώνυμο *", 194),
        ("Forms.TextBox.1", "txtDisplayName", None, 190),
        ("Forms.Label.1", "lblRoom", "Θάλαμος", 238),
        ("Forms.ComboBox.1", "cboRoom", None, 234),
        ("Forms.Label.1", "lblInfectious", "Λοιμώδης", 282),
        ("Forms.CheckBox.1", "chkInfectious", "Ναι", 278),
        ("Forms.Label.1", "lblStatus", "Κατάσταση", 326),
        ("Forms.ComboBox.1", "cboStatus", None, 322),
        ("Forms.Label.1", "lblInfo", "", 370),
        ("Forms.Label.1", "lblRequired", "* Υποχρεωτικό πεδίο", 414),
        ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 452),
        ("Forms.CommandButton.1", "cmdSave", "Έλεγχος στοιχείων", 452),
    )

    for prog_id, name, caption, top in controls:
        control = designer.Controls.Add(prog_id, name, True)
        if caption is not None:
            control.Caption = caption
        position_control(control, 24, top)

    form.CodeModule.AddFromString(PATIENT_FORM_CODE)
