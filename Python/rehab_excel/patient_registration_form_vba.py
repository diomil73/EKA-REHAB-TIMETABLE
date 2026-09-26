from __future__ import annotations

PATIENT_FORM_NAME = "frmNewPatient"

PATIENT_FORM_CODE = '''Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Νέος ασθενής"
        .Width = 430
        .Height = 430
        .StartUpPosition = 1
        .BackColor = RGB(245, 247, 250)
    End With

    StyleTitle lblTitle, "Εγγραφή νέου ασθενή", 20
    StyleLabel lblPatientID, "Patient ID *", 64
    StyleTextBox txtPatientID, 60
    StyleLabel lblDisplayName, "Ονοματεπώνυμο *", 108
    StyleTextBox txtDisplayName, 104
    StyleLabel lblRoom, "Θάλαμος", 152
    StyleComboBox cboRoom, 148
    StyleLabel lblInfectious, "Λοιμώδης", 196
    StyleCheckBox chkInfectious, "Ναι", 192
    StyleLabel lblStatus, "Κατάσταση", 240
    StyleComboBox cboStatus, 236

    With lblRequired
        .Caption = "* Υποχρεωτικό πεδίο"
        .Left = 36
        .Top = 286
        .Width = 180
        .Height = 18
        .Font.Name = "Calibri"
        .Font.Size = 9
        .ForeColor = RGB(100, 116, 139)
        .BackStyle = 0
    End With

    With cmdSave
        .Caption = "Έλεγχος στοιχείων"
        .Left = 178
        .Top = 325
        .Width = 150
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Default = True
    End With

    With cmdCancel
        .Caption = "Ακύρωση"
        .Left = 58
        .Top = 325
        .Width = 105
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Cancel = True
    End With

    LoadSettingsValues
    txtPatientID.SetFocus
End Sub

Private Sub StyleTitle(ByVal control As MSForms.Label, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 36
        .Top = topPosition
        .Width = 350
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
        .Left = 36
        .Top = topPosition
        .Width = 120
        .Height = 20
        .Font.Name = "Calibri"
        .Font.Size = 10
        .ForeColor = RGB(51, 65, 85)
        .BackStyle = 0
    End With
End Sub

Private Sub StyleTextBox(ByVal control As MSForms.TextBox, ByVal topPosition As Single)
    With control
        .Left = 165
        .Top = topPosition
        .Width = 220
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub

Private Sub StyleComboBox(ByVal control As MSForms.ComboBox, ByVal topPosition As Single)
    With control
        .Left = 165
        .Top = topPosition
        .Width = 220
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Style = 2
    End With
End Sub

Private Sub StyleCheckBox(ByVal control As MSForms.CheckBox, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 165
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

Private Function ValidateForm() As Boolean
    If Len(Trim$(txtPatientID.Text)) = 0 Then
        MsgBox "Το Patient ID είναι υποχρεωτικό.", vbExclamation, "Έλεγχος στοιχείων"
        txtPatientID.SetFocus
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

    If Not ValidateForm() Then Exit Sub

    message = "Patient ID: " & Trim$(txtPatientID.Text) & vbCrLf & _
              "Ονοματεπώνυμο: " & Trim$(txtDisplayName.Text) & vbCrLf & _
              "Θάλαμος: " & IIf(Len(cboRoom.Value) > 0, cboRoom.Value, "-") & vbCrLf & _
              "Λοιμώδης: " & IIf(chkInfectious.Value, "Ναι", "Όχι") & vbCrLf & _
              "Κατάσταση: " & IIf(Len(cboStatus.Value) > 0, cboStatus.Value, "-") & vbCrLf & vbCrLf & _
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
        ("Forms.Label.1", "lblTitle", "Εγγραφή νέου ασθενή", 20),
        ("Forms.Label.1", "lblPatientID", "Patient ID *", 64),
        ("Forms.TextBox.1", "txtPatientID", None, 60),
        ("Forms.Label.1", "lblDisplayName", "Ονοματεπώνυμο *", 108),
        ("Forms.TextBox.1", "txtDisplayName", None, 104),
        ("Forms.Label.1", "lblRoom", "Θάλαμος", 152),
        ("Forms.ComboBox.1", "cboRoom", None, 148),
        ("Forms.Label.1", "lblInfectious", "Λοιμώδης", 196),
        ("Forms.CheckBox.1", "chkInfectious", "Ναι", 192),
        ("Forms.Label.1", "lblStatus", "Κατάσταση", 240),
        ("Forms.ComboBox.1", "cboStatus", None, 236),
        ("Forms.Label.1", "lblRequired", "* Υποχρεωτικό πεδίο", 286),
        ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 325),
        ("Forms.CommandButton.1", "cmdSave", "Έλεγχος στοιχείων", 325),
    )

    for prog_id, name, caption, top in controls:
        control = designer.Controls.Add(prog_id, name, True)
        if caption is not None:
            control.Caption = caption
        position_control(control, 24, top)

    form.CodeModule.AddFromString(PATIENT_FORM_CODE)
