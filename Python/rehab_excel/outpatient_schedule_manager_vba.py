from __future__ import annotations


FORM_NAME = "frmOutpatientScheduleManager"
MODULE_NAME = "modOutpatientScheduleManager"
MACRO_NAME = "ShowOutpatientScheduleManager"
REGISTRATION_MENU_NAME = "frmRegistrationMenu"


MODULE_CODE = r'''Option Explicit

Public Sub ShowOutpatientScheduleManager()
    Load frmOutpatientScheduleManager
    frmOutpatientScheduleManager.Show
End Sub
'''


FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Προγράμματα εξωτερικών ασθενών"
    Me.Width = 525
    Me.Height = 430

    lblTitle.Caption = "Προγράμματα εξωτερικών ασθενών"
    lblTitle.Width = 465
    lblTitle.Height = 24

    lblInfo.Caption = "Επιλέξτε μία εγγραφή για επεξεργασία ή δημιουργήστε νέο πρόγραμμα."
    lblInfo.Width = 465
    lblInfo.Height = 30

    lblPatient.Caption = "Ασθενής"
    lblTreatment.Caption = "Θεραπεία"
    lblTime.Caption = "Ώρα"
    lblDays.Caption = "Ημέρες"
    lblTherapist.Caption = "Θεραπευτής"

    With lstSchedules
        .Width = 465
        .Height = 232
        .ColumnCount = 6
        .ColumnWidths = "0 pt;150 pt;95 pt;55 pt;85 pt;110 pt"
        .IntegralHeight = False
    End With

    cmdNew.Caption = "Νέο πρόγραμμα"
    cmdNew.Width = 135
    cmdNew.Height = 30

    cmdEdit.Caption = "Επεξεργασία"
    cmdEdit.Width = 135
    cmdEdit.Height = 30

    cmdClose.Caption = "Κλείσιμο"
    cmdClose.Width = 135
    cmdClose.Height = 30

    RefreshScheduleList
End Sub

Private Sub RefreshScheduleList()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim patientId As String
    Dim patientName As String
    Dim treatment As String
    Dim timeText As String
    Dim dayPattern As String
    Dim therapistName As String

    lstSchedules.Clear
    Set ws = ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 2 To lastRow
        patientId = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
        patientName = Trim$(CStr(ws.Cells(rowIndex, 2).Value))
        treatment = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
        dayPattern = Trim$(CStr(ws.Cells(rowIndex, 5).Value))
        therapistName = Trim$(CStr(ws.Cells(rowIndex, 6).Value))

        If Len(patientId) > 0 And Len(patientName) > 0 And Len(treatment) > 0 And Len(dayPattern) > 0 Then
            If IsDate(ws.Cells(rowIndex, 4).Value) Then
                timeText = Format$(ws.Cells(rowIndex, 4).Value, "hh:mm")
            Else
                timeText = Trim$(CStr(ws.Cells(rowIndex, 4).Text))
            End If

            If Len(timeText) > 0 Then
                lstSchedules.AddItem CStr(rowIndex)
                lstSchedules.List(lstSchedules.ListCount - 1, 1) = patientName
                lstSchedules.List(lstSchedules.ListCount - 1, 2) = treatment
                lstSchedules.List(lstSchedules.ListCount - 1, 3) = timeText
                lstSchedules.List(lstSchedules.ListCount - 1, 4) = dayPattern
                lstSchedules.List(lstSchedules.ListCount - 1, 5) = therapistName
            End If
        End If
    Next rowIndex

    If lstSchedules.ListCount = 0 Then
        lblInfo.Caption = "Δεν υπάρχουν ακόμη καταχωρημένα προγράμματα εξωτερικών ασθενών."
        cmdEdit.Enabled = False
    Else
        lblInfo.Caption = "Επιλέξτε μία εγγραφή για επεξεργασία ή δημιουργήστε νέο πρόγραμμα."
        cmdEdit.Enabled = True
    End If
End Sub

Private Function SelectedScheduleRow() As Long
    If lstSchedules.ListIndex < 0 Then
        MsgBox "Επιλέξτε πρώτα το πρόγραμμα που θέλετε να επεξεργαστείτε.", _
               vbInformation, "Προγράμματα εξωτερικών ασθενών"
        Exit Function
    End If

    SelectedScheduleRow = CLng(lstSchedules.List(lstSchedules.ListIndex, 0))
End Function

Private Sub cmdNew_Click()
    Unload Me
    ShowOutpatientScheduleForm
End Sub

Private Sub cmdEdit_Click()
    Dim targetRow As Long
    targetRow = SelectedScheduleRow()
    If targetRow = 0 Then Exit Sub

    Unload Me
    EditOutpatientScheduleByRow targetRow
End Sub

Private Sub lstSchedules_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdEdit_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'''


REGISTRATION_MENU_OUTPATIENT_CLICK = r'''Private Sub cmdOutpatientSchedule_Click()
    On Error GoTo MissingManager
    Unload Me
    ShowOutpatientScheduleManager
    Exit Sub
MissingManager:
    MsgBox "Η διαχείριση προγραμμάτων εξωτερικών ασθενών δεν είναι διαθέσιμη σε αυτό το αρχείο.", _
           vbExclamation, "Προγράμματα εξωτερικών ασθενών"
End Sub'''


def _component_if_present(vbproject, name: str):
    try:
        return vbproject.VBComponents(name)
    except Exception:
        return None


def _component_code(component) -> str:
    module = component.CodeModule
    count = int(getattr(module, "CountOfLines", 0) or 0)
    if count <= 0:
        return ""
    return str(module.Lines(1, count))


def _replace_component_code(component, code: str) -> None:
    module = component.CodeModule
    count = int(getattr(module, "CountOfLines", 0) or 0)
    if count > 0:
        module.DeleteLines(1, count)
    module.AddFromString(code)


def _ensure_form_control(designer, prog_id: str, name: str):
    try:
        return designer.Controls(name)
    except Exception:
        return designer.Controls.Add(prog_id, name, True)


def _replace_vba_procedure(
    code: str,
    *,
    start_signature: str,
    end_statement: str,
    replacement: str,
) -> str:
    normalized = code.replace("\r\n", "\n").replace("\r", "\n")
    start = normalized.find(start_signature)
    if start < 0:
        raise RuntimeError(f"VBA procedure missing: {start_signature}")
    if normalized.find(start_signature, start + len(start_signature)) >= 0:
        raise RuntimeError(f"VBA procedure duplicated: {start_signature}")
    end_marker = "\n" + end_statement
    end = normalized.find(end_marker, start + len(start_signature))
    if end < 0:
        raise RuntimeError(f"VBA procedure terminator missing: {start_signature}")
    end += len(end_marker)
    return normalized[:start] + replacement + normalized[end:]


def patch_registration_menu_for_outpatient_manager(code: str) -> str:
    patched = code.replace("\r\n", "\n").replace("\r", "\n")
    patched = patched.replace(
        'StyleMenuButton cmdOutpatientSchedule, "Πρόγραμμα εξωτερικού ασθενή",',
        'StyleMenuButton cmdOutpatientSchedule, "Προγράμματα εξωτερικών ασθενών",',
    )

    if "StyleMenuButton cmdOutpatientSchedule" not in patched:
        anchor = 'StyleMenuButton cmdStudent, "Νέος φοιτητής", 158'
        if anchor in patched:
            patched = patched.replace(
                anchor,
                anchor
                + '\n    StyleMenuButton cmdOutpatientSchedule, "Προγράμματα εξωτερικών ασθενών", 203',
                1,
            )

    # Older menu builds placed Close where the outpatient action now belongs.
    # Move it down when that exact legacy layout is present.
    close_start = patched.find("With cmdClose")
    if close_start >= 0:
        close_end = patched.find("End With", close_start)
        if close_end >= 0:
            close_block = patched[close_start:close_end]
            new_close_block = close_block.replace(".Top = 215", ".Top = 260")
            patched = patched[:close_start] + new_close_block + patched[close_end:]

    if "Private Sub cmdOutpatientSchedule_Click()" in patched:
        patched = _replace_vba_procedure(
            patched,
            start_signature="Private Sub cmdOutpatientSchedule_Click()",
            end_statement="End Sub",
            replacement=REGISTRATION_MENU_OUTPATIENT_CLICK,
        )
    else:
        patched = patched.rstrip() + "\n\n" + REGISTRATION_MENU_OUTPATIENT_CLICK + "\n"

    return patched


def _ensure_registration_menu_entry(registration_menu) -> None:
    designer = registration_menu.Designer
    button = _ensure_form_control(
        designer,
        "Forms.CommandButton.1",
        "cmdOutpatientSchedule",
    )
    button.Caption = "Προγράμματα εξωτερικών ασθενών"
    button.Left = 45
    button.Top = 203
    button.Width = 230
    button.Height = 34
    try:
        button.Font.Name = "Calibri"
        button.Font.Size = 11
        button.Font.Bold = True
        button.TakeFocusOnClick = False
    except Exception:
        pass

    try:
        close_button = designer.Controls("cmdClose")
        close_button.Top = 260
    except Exception:
        pass

    registration_code = _component_code(registration_menu)
    patched_registration = patch_registration_menu_for_outpatient_manager(registration_code)
    if patched_registration != registration_code.replace("\r\n", "\n").replace("\r", "\n"):
        _replace_component_code(registration_menu, patched_registration)


def install_outpatient_schedule_manager(vbproject) -> None:
    """Install the end-user outpatient schedule selection UI in place."""

    module = _component_if_present(vbproject, MODULE_NAME)
    if module is None:
        module = vbproject.VBComponents.Add(1)
        module.Name = MODULE_NAME
    _replace_component_code(module, MODULE_CODE)

    form = _component_if_present(vbproject, FORM_NAME)
    if form is None:
        form = vbproject.VBComponents.Add(3)
        form.Name = FORM_NAME

    designer = form.Designer
    designer.Caption = "Προγράμματα εξωτερικών ασθενών"

    controls = (
        ("Forms.Label.1", "lblTitle", 18, 12),
        ("Forms.Label.1", "lblInfo", 18, 40),
        ("Forms.Label.1", "lblPatient", 24, 78),
        ("Forms.Label.1", "lblTreatment", 174, 78),
        ("Forms.Label.1", "lblTime", 269, 78),
        ("Forms.Label.1", "lblDays", 324, 78),
        ("Forms.Label.1", "lblTherapist", 409, 78),
        ("Forms.ListBox.1", "lstSchedules", 18, 96),
        ("Forms.CommandButton.1", "cmdNew", 18, 342),
        ("Forms.CommandButton.1", "cmdEdit", 168, 342),
        ("Forms.CommandButton.1", "cmdClose", 348, 342),
    )
    for prog_id, name, left, top in controls:
        control = _ensure_form_control(designer, prog_id, name)
        control.Left = left
        control.Top = top

    _replace_component_code(form, FORM_CODE)

    registration_menu = _component_if_present(vbproject, REGISTRATION_MENU_NAME)
    if registration_menu is None:
        raise RuntimeError(
            "Registration menu is missing; outpatient management would have no application entry point"
        )
    _ensure_registration_menu_entry(registration_menu)

    installed_module = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)
    installed_form = form.CodeModule.Lines(1, form.CodeModule.CountOfLines)
    installed_registration = _component_code(registration_menu)
    if "Public Sub ShowOutpatientScheduleManager()" not in installed_module:
        raise RuntimeError("Outpatient schedule manager macro was not installed")
    if "EditOutpatientScheduleByRow targetRow" not in installed_form:
        raise RuntimeError("Outpatient schedule manager edit action was not installed")
    if "ShowOutpatientScheduleManager" not in installed_registration:
        raise RuntimeError("Registration menu does not route to outpatient schedule manager")
