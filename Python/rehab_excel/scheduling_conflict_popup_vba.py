from __future__ import annotations


SCHEDULING_CONFLICT_FORM_NAME = "frmSchedulingConflict"
SCHEDULING_CONFLICT_MODULE_NAME = "modSchedulingConflict"

SCHEDULING_CONFLICT_MODULE_CODE = r'''Option Explicit

Public Function ConfirmTherapistDoubleBooking( _
    ByVal therapistName As String, _
    ByVal timeText As String, _
    ByVal existingPatientName As String, _
    ByVal newPatientName As String) As Boolean

    Dim dlg As frmSchedulingConflict
    Set dlg = New frmSchedulingConflict

    dlg.ConfigureConflict therapistName, timeText, existingPatientName, newPatientName
    Beep
    dlg.Show vbModal

    ConfirmTherapistDoubleBooking = dlg.Accepted
    Unload dlg
    Set dlg = Nothing
End Function
'''

SCHEDULING_CONFLICT_FORM_CODE = r'''Option Explicit

Public Accepted As Boolean

Private Sub UserForm_Initialize()
    Accepted = False

    With Me
        .Caption = "ΠΡΟΣΟΧΗ - ΔΙΠΛΟΚΡΑΤΗΣΗ"
        .Width = 520
        .Height = 330
        .StartUpPosition = 1
        .BackColor = RGB(255, 248, 235)
    End With

    With lblTitle
        .Caption = "ΠΡΟΣΟΧΗ: Ο θεραπευτής έχει δύο ασθενείς ταυτόχρονα"
        .Left = 28
        .Top = 20
        .Width = 455
        .Height = 42
        .WordWrap = True
        .TextAlign = 2
        .Font.Name = "Calibri"
        .Font.Size = 13
        .Font.Bold = True
        .ForeColor = RGB(153, 27, 27)
        .BackStyle = 0
    End With

    With lblMessage
        .Left = 35
        .Top = 78
        .Width = 440
        .Height = 95
        .WordWrap = True
        .Font.Name = "Calibri"
        .Font.Size = 11
        .ForeColor = RGB(31, 41, 55)
        .BackStyle = 0
    End With

    With lblQuestion
        .Caption = "Θέλετε να το δεχτείτε συνειδητά ή να επιστρέψετε για αλλαγή;"
        .Left = 35
        .Top = 182
        .Width = 440
        .Height = 38
        .WordWrap = True
        .TextAlign = 2
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .BackStyle = 0
    End With

    With cmdChange
        .Caption = "Αλλαγή"
        .Left = 90
        .Top = 238
        .Width = 145
        .Height = 38
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Default = True
        .Cancel = True
    End With

    With cmdAccept
        .Caption = "Δεκτό"
        .Left = 275
        .Top = 238
        .Width = 145
        .Height = 38
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
    End With
End Sub

Public Sub ConfigureConflict( _
    ByVal therapistName As String, _
    ByVal timeText As String, _
    ByVal existingPatientName As String, _
    ByVal newPatientName As String)

    lblMessage.Caption = _
        "Η/Ο " & therapistName & " έχει ήδη τον/την " & existingPatientName & _
        " στις " & timeText & "." & vbCrLf & vbCrLf & _
        "Η νέα καταχώρηση είναι για τον/την " & newPatientName & _
        " στην ίδια ώρα."
End Sub

Private Sub cmdAccept_Click()
    Accepted = True
    Me.Hide
End Sub

Private Sub cmdChange_Click()
    Accepted = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Accepted = False
    End If
End Sub
'''


def install_scheduling_conflict_popup(vbproject, *, position_control) -> None:
    for name in (SCHEDULING_CONFLICT_FORM_NAME, SCHEDULING_CONFLICT_MODULE_NAME):
        try:
            existing = vbproject.VBComponents(name)
        except Exception:
            existing = None
        if existing is not None:
            vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = SCHEDULING_CONFLICT_MODULE_NAME
    module.CodeModule.AddFromString(SCHEDULING_CONFLICT_MODULE_CODE)

    form = vbproject.VBComponents.Add(3)
    form.Name = SCHEDULING_CONFLICT_FORM_NAME
    designer = form.Designer
    designer.Caption = "ΠΡΟΣΟΧΗ - ΔΙΠΛΟΚΡΑΤΗΣΗ"

    controls = (
        ("Forms.Label.1", "lblTitle", "", 20),
        ("Forms.Label.1", "lblMessage", "", 78),
        ("Forms.Label.1", "lblQuestion", "", 182),
        ("Forms.CommandButton.1", "cmdChange", "Αλλαγή", 238),
        ("Forms.CommandButton.1", "cmdAccept", "Δεκτό", 238),
    )
    for prog_id, name, caption, top in controls:
        control = designer.Controls.Add(prog_id, name, True)
        if caption:
            control.Caption = caption
        position_control(control, 24, top)

    form.CodeModule.AddFromString(SCHEDULING_CONFLICT_FORM_CODE)
