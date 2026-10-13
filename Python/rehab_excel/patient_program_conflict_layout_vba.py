from __future__ import annotations


CONFLICT_FORM_NAME = "frmPatientProgramConflict"


CONFLICT_FORM_RUNTIME_CODE = r'''Option Explicit

Public ContinueSelected As Boolean

Private Sub ApplyRuntimeLayout()
    On Error Resume Next

    Me.Width = 620
    Me.Height = 390

    lblConflict.Left = 18
    lblConflict.Top = 18
    lblConflict.Width = 570
    lblConflict.Height = 285
    lblConflict.WordWrap = True

    cmdContinue.Left = 320
    cmdContinue.Top = 320
    cmdContinue.Width = 125
    cmdContinue.Height = 32

    cmdReturn.Left = 455
    cmdReturn.Top = 320
    cmdReturn.Width = 125
    cmdReturn.Height = 32

    On Error GoTo 0
End Sub

Public Sub Configure(ByVal conflictText As String)
    ContinueSelected = False
    ApplyRuntimeLayout
    Me.Caption = "Σύγκρουση στο πρόγραμμα"
    lblConflict.Caption = "Εντοπίστηκε η παρακάτω σύγκρουση:" & vbCrLf & vbCrLf & _
                          conflictText & vbCrLf & vbCrLf & _
                          "Θέλετε να συνεχίσετε την αποθήκευση ή να επιστρέψετε για διόρθωση;"
End Sub

Private Sub cmdContinue_Click()
    ContinueSelected = True
    Me.Hide
End Sub

Private Sub cmdReturn_Click()
    ContinueSelected = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ContinueSelected = False
End Sub
'''


def install_conflict_runtime_layout(vbproject) -> None:
    """Replace conflict-form code with runtime geometry that survives VBE sizing quirks."""

    form = vbproject.VBComponents(CONFLICT_FORM_NAME)
    code_module = form.CodeModule
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    if count:
        code_module.DeleteLines(1, count)
    code_module.AddFromString(CONFLICT_FORM_RUNTIME_CODE)
