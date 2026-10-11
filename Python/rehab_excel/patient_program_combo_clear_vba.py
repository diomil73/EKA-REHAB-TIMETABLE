from __future__ import annotations

PATIENT_PROGRAM_FORM_NAME = "frmPatientScheduling"

COMBO_CLEAR_HANDLER_CODE = r'''
Private Sub ClearProgramComboSelection(ByVal combo As Object, ByRef KeyCode As MSForms.ReturnInteger)
    If Len(mCurrentTreatment) = 0 Then Exit Sub

    mLoading = True
    combo.ListIndex = -1
    mLoading = False

    SaveCurrentDraft
    UpdateDraftStatus
    KeyCode = 0
End Sub

Private Sub cboDayPattern_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        ClearProgramComboSelection cboDayPattern, KeyCode
    End If
End Sub

Private Sub cboTime_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        ClearProgramComboSelection cboTime, KeyCode
    End If
End Sub

Private Sub cboProvider_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        ClearProgramComboSelection cboProvider, KeyCode
    End If
End Sub
'''


def install_combo_clear_handlers(vbproject) -> None:
    """Append keyboard-clear handlers to the freshly installed patient-program form.

    The main form installer rewrites the whole form code module on every preview
    build, so this patch is intentionally applied immediately afterwards. Delete
    or Backspace clears the active combo selection and therefore removes the
    corresponding draft constraint, allowing automatic completion to fill it
    again.
    """

    form = vbproject.VBComponents(PATIENT_PROGRAM_FORM_NAME)
    form.CodeModule.AddFromString(COMBO_CLEAR_HANDLER_CODE)
