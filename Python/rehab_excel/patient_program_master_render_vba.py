from __future__ import annotations


MASTER_RENDER_MODULE_NAME = "PatientProgramMasterRender"
PATIENT_PROGRAM_RUNTIME_NAME = "PatientProgramRuntime"
PATIENT_PROGRAM_FORM_NAME = "frmPatientScheduling"
PATIENT_EDIT_FORM_NAME = "frmEditPatient"
CONFLICT_FORM_NAME = "frmPatientProgramConflict"


MASTER_RENDER_MODULE_CODE = r'''Option Explicit

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim colIndex As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For colIndex = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, colIndex).Value)), Trim$(headerName), vbTextCompare) = 0 Then
            FindHeaderColumn = colIndex
            Exit Function
        End If
    Next colIndex
End Function

Private Function FindPatientRowById(ByVal ws As Worksheet, ByVal patientId As String) As Long
    Dim idCol As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    idCol = FindHeaderColumn(ws, "PatientID")
    If idCol = 0 Then idCol = 1
    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, idCol).Value)), Trim$(patientId), vbTextCompare) = 0 Then
            FindPatientRowById = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function ProviderHeaderName(ByVal treatment As String) As String
    Select Case LCase$(Trim$(treatment))
        Case LCase$("ΦΘ"), LCase$("Ρομποτικό"), LCase$("Εργο"), LCase$("Λογο")
            ProviderHeaderName = treatment & "_Θεραπευτής"
        Case LCase$("Ψυχ")
            ProviderHeaderName = "Ψυχ_Ψυχολόγος"
    End Select
End Function

Private Function TimeText(ByVal rawValue As Variant) As String
    If IsEmpty(rawValue) Or Len(Trim$(CStr(rawValue))) = 0 Then Exit Function
    If IsDate(rawValue) Or IsNumeric(rawValue) Then
        TimeText = Format$(rawValue, "hh:mm")
    Else
        TimeText = Trim$(CStr(rawValue))
    End If
End Function

Private Function BuildTreatmentDisplay(ByVal planner As Worksheet, ByVal rowIndex As Long, ByVal treatment As String) As String
    Dim timeCol As Long
    Dim daysCol As Long
    Dim providerCol As Long
    Dim providerHeader As String
    Dim result As String
    Dim part As String

    timeCol = FindHeaderColumn(planner, treatment & "_Ώρα")
    daysCol = FindHeaderColumn(planner, treatment & "_Ημέρες")
    providerHeader = ProviderHeaderName(treatment)
    If Len(providerHeader) > 0 Then providerCol = FindHeaderColumn(planner, providerHeader)

    If timeCol > 0 Then result = TimeText(planner.Cells(rowIndex, timeCol).Value)

    If daysCol > 0 Then
        part = Trim$(CStr(planner.Cells(rowIndex, daysCol).Value))
        If Len(part) > 0 Then
            If Len(result) > 0 Then result = result & vbLf
            result = result & part
        End If
    End If

    If providerCol > 0 Then
        part = Trim$(CStr(planner.Cells(rowIndex, providerCol).Value))
        If Len(part) > 0 Then
            If Len(result) > 0 Then result = result & vbLf
            result = result & part
        End If
    End If

    BuildTreatmentDisplay = result
End Function

Private Function FindMasterHeaderRow(ByVal ws As Worksheet) As Long
    Dim rowIndex As Long
    For rowIndex = 1 To 10
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 2).Value)), "Θάλαμος", vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(ws.Cells(rowIndex, 3).Value)), "Ασθενής", vbTextCompare) = 0 Then
            FindMasterHeaderRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FirstLine(ByVal valueText As String) As String
    Dim normalized As String
    Dim parts() As String

    normalized = Trim$(Replace(valueText, vbCrLf, vbLf))
    If Len(normalized) = 0 Then
        FirstLine = ""
        Exit Function
    End If

    parts = Split(normalized, vbLf)
    FirstLine = Trim$(parts(0))
End Function

Private Function FindMasterPatientRow(ByVal ws As Worksheet, ByVal roomValue As String, ByVal displayName As String) As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    headerRow = FindMasterHeaderRow(ws)
    If headerRow = 0 Then Exit Function
    lastRow = Application.Max(ws.Cells(ws.Rows.Count, 2).End(xlUp).Row, ws.Cells(ws.Rows.Count, 3).End(xlUp).Row)

    For rowIndex = headerRow + 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 2).Value)), Trim$(roomValue), vbTextCompare) = 0 _
           And StrComp(FirstLine(CStr(ws.Cells(rowIndex, 3).Value)), Trim$(displayName), vbTextCompare) = 0 Then
            FindMasterPatientRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub WriteTreatmentCell(ByVal targetCell As Range, ByVal planner As Worksheet, ByVal plannerRow As Long, ByVal treatment As String)
    With targetCell
        .NumberFormat = "General"
        .Value = BuildTreatmentDisplay(planner, plannerRow, treatment)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub RefreshMasterPatientProgram(ByVal patientId As String)
    Dim patients As Worksheet
    Dim planner As Worksheet
    Dim master As Worksheet
    Dim patientRow As Long
    Dim plannerRow As Long
    Dim masterRow As Long
    Dim displayName As String
    Dim roomValue As String
    Dim headerRow As Long
    Dim renderStage As String
    Dim errorText As String

    On Error GoTo RenderError

    renderStage = "opening PATIENTS"
    Set patients = ThisWorkbook.Worksheets("PATIENTS")
    renderStage = "opening PATIENT_PLANNER"
    Set planner = ThisWorkbook.Worksheets("PATIENT_PLANNER")
    renderStage = "opening MASTER_SCHEDULE"
    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")

    renderStage = "locating patient rows"
    patientRow = FindPatientRowById(patients, patientId)
    plannerRow = FindPatientRowById(planner, patientId)
    If patientRow = 0 Or plannerRow = 0 Then Exit Sub

    renderStage = "reading patient identity"
    displayName = Trim$(CStr(patients.Cells(patientRow, 3).Value))
    roomValue = Trim$(CStr(patients.Cells(patientRow, 2).Value))

    renderStage = "locating MASTER row"
    masterRow = FindMasterPatientRow(master, roomValue, displayName)
    If masterRow = 0 Then Exit Sub

    renderStage = "writing treatment cells"
    WriteTreatmentCell master.Cells(masterRow, 4), planner, plannerRow, "ΦΘ"
    WriteTreatmentCell master.Cells(masterRow, 5), planner, plannerRow, "Ρομποτικό"
    WriteTreatmentCell master.Cells(masterRow, 6), planner, plannerRow, "Πισίνα"
    WriteTreatmentCell master.Cells(masterRow, 7), planner, plannerRow, "Ανακλινόμενο"
    WriteTreatmentCell master.Cells(masterRow, 8), planner, plannerRow, "Εργο"
    WriteTreatmentCell master.Cells(masterRow, 9), planner, plannerRow, "Λογο"
    WriteTreatmentCell master.Cells(masterRow, 10), planner, plannerRow, "ΕΦΑ"

    renderStage = "writing psychology cell"
    headerRow = FindMasterHeaderRow(master)
    If headerRow > 0 Then
        If StrComp(Trim$(CStr(master.Cells(headerRow, 11).Value)), "ΨΥΧΟΛΟΓΟΙ", vbTextCompare) = 0 Then
            WriteTreatmentCell master.Cells(masterRow, 11), planner, plannerRow, "Ψυχ"
        End If
    End If
    Exit Sub

RenderError:
    errorText = Err.Description
    Err.Raise vbObjectError + 2940, "RefreshMasterPatientProgram", _
        "MASTER refresh failed at " & renderStage & ": " & errorText
End Sub
'''


CONFLICT_FORM_CODE = r'''Option Explicit

Public ContinueSelected As Boolean

Public Sub Configure(ByVal conflictText As String)
    ContinueSelected = False
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


def _get_component(vbproject, name: str):
    try:
        return vbproject.VBComponents(name)
    except Exception:
        return None


def _replace_component_code(component, code: str) -> None:
    code_module = component.CodeModule
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    if count:
        code_module.DeleteLines(1, count)
    code_module.AddFromString(code)


def _normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _write_code_module(code_module, text: str) -> None:
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    if count:
        code_module.DeleteLines(1, count)
    code_module.AddFromString(text)


def _patch_code_module(code_module, old: str, new: str, *, description: str) -> None:
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    raw_text = code_module.Lines(1, count) if count else ""
    text = _normalize_newlines(raw_text)
    old_normalized = _normalize_newlines(old)
    new_normalized = _normalize_newlines(new)

    if new_normalized in text:
        return
    if old_normalized not in text:
        raise RuntimeError(f"MASTER render patch marker missing: {description}")

    _write_code_module(code_module, text.replace(old_normalized, new_normalized, 1))


def _ensure_conflict_control(designer, prog_id: str, name: str, caption: str, left: float, top: float, width: float, height: float):
    try:
        control = designer.Controls(name)
    except Exception:
        control = designer.Controls.Add(prog_id, name, True)
    control.Caption = caption
    try:
        control.Left = left
        control.Top = top
    except Exception:
        pass
    try:
        control.Width = width
        control.Height = height
    except Exception:
        pass
    return control


def _install_conflict_form(vbproject) -> None:
    form = _get_component(vbproject, CONFLICT_FORM_NAME)
    if form is None:
        form = vbproject.VBComponents.Add(3)
        form.Name = CONFLICT_FORM_NAME
    designer = form.Designer
    designer.Caption = "Σύγκρουση στο πρόγραμμα"
    try:
        designer.Width = 620
        designer.Height = 330
    except Exception:
        pass

    label = _ensure_conflict_control(
        designer, "Forms.Label.1", "lblConflict", "", 18, 18, 570, 220
    )
    try:
        label.WordWrap = True
    except Exception:
        pass

    continue_button = _ensure_conflict_control(
        designer, "Forms.CommandButton.1", "cmdContinue", "Συνέχεια", 320, 252, 125, 32
    )
    return_button = _ensure_conflict_control(
        designer, "Forms.CommandButton.1", "cmdReturn", "Επιστροφή", 455, 252, 125, 32
    )
    try:
        continue_button.Default = False
        return_button.Cancel = True
    except Exception:
        pass

    _replace_component_code(form, CONFLICT_FORM_CODE)


def _patch_runtime_validation(code_module) -> None:
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    raw_text = code_module.Lines(1, count) if count else ""
    text = _normalize_newlines(raw_text)

    validation_marker = "Public Function ValidatePatientProgram(ByVal patientId As String, ByVal requestArguments As String) As String"
    if validation_marker in text:
        return

    procedure_marker = "Public Function CompletePatientProgram(ByVal patientId As String, ByVal requestArguments As String) As String"
    procedure_start = text.find(procedure_marker)
    if procedure_start < 0:
        raise RuntimeError("MASTER render patch marker missing: patient-program CompletePatientProgram")

    procedure_end = text.find("\nEnd Function", procedure_start)
    if procedure_end < 0:
        raise RuntimeError("MASTER render patch marker missing: patient-program CompletePatientProgram end")
    procedure_end += len("\nEnd Function")

    validation_block = r'''

Public Function ValidatePatientProgram(ByVal patientId As String, ByVal requestArguments As String) As String
    ValidatePatientProgram = RunProgramBridge( _
        "validate-vba " & QuoteArg(ThisWorkbook.FullName) & _
        " --patient-id " & QuoteArg(patientId) & requestArguments _
    )
End Function'''

    patched_text = text[:procedure_end] + validation_block + text[procedure_end:]
    _write_code_module(code_module, patched_text)


def _patch_program_conflict_validation(code_module) -> None:
    helper_marker = "Private Sub ApplyAutomaticCompletion(ByVal response As String)"
    helper_block = r'''Private Function ConflictTextFromValidation(ByVal response As String) As String
    Dim lines() As String
    Dim fields() As String
    Dim i As Long

    lines = Split(Replace(response, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) = 0 Then GoTo NextLine
        fields = Split(lines(i), vbTab)
        If fields(0) = "ERROR" Then Err.Raise vbObjectError + 2922, , fields(1)
        If fields(0) = "CONFLICT" And UBound(fields) >= 6 Then
            If Len(ConflictTextFromValidation) > 0 Then _
                ConflictTextFromValidation = ConflictTextFromValidation & vbCrLf
            ConflictTextFromValidation = ConflictTextFromValidation & fields(6)
        End If
NextLine:
    Next i
End Function

Private Function CurrentConflictText() As String
    Dim requestArguments As String
    requestArguments = BuildAutomaticRequestArguments()
    If Len(Trim$(requestArguments)) = 0 Then Exit Function
    CurrentConflictText = ConflictTextFromValidation( _
        ValidatePatientProgram(mPatientId, requestArguments) _
    )
End Function

Private Function ConfirmConflictOverride(ByVal conflictText As String) As Boolean
    Load frmPatientProgramConflict
    frmPatientProgramConflict.Configure conflictText
    frmPatientProgramConflict.Show vbModal
    ConfirmConflictOverride = frmPatientProgramConflict.ContinueSelected
    Unload frmPatientProgramConflict
End Function

Private Sub ApplyAutomaticCompletion(ByVal response As String)'''
    _patch_code_module(
        code_module,
        helper_marker,
        helper_block,
        description="patient-program conflict helpers",
    )

    auto_old = '''Private Sub cmdAutoComplete_Click()
    Dim requestArguments As String
    Dim response As String

    On Error GoTo AutoError'''
    auto_new = '''Private Sub cmdAutoComplete_Click()
    Dim requestArguments As String
    Dim response As String
    Dim conflictText As String

    On Error GoTo AutoError'''
    _patch_code_module(
        code_module,
        auto_old,
        auto_new,
        description="patient-program auto conflict declaration",
    )

    auto_validation_old = '''    lblStatus.Caption = "Υπολογίζονται τα κενά στοιχεία του προγράμματος..."
    DoEvents
    response = CompletePatientProgram(mPatientId, requestArguments)'''
    auto_validation_new = '''    conflictText = CurrentConflictText()
    If Len(conflictText) > 0 Then
        lblStatus.Caption = "Υπάρχει χρονική σύγκρουση στο πρόγραμμα."
        MsgBox conflictText & vbCrLf & vbCrLf & _
               "Διόρθωσε τη σύγκρουση και ξαναπάτησε Αυτόματη συμπλήρωση.", _
               vbExclamation, "Σύγκρουση στο πρόγραμμα"
        Exit Sub
    End If

    lblStatus.Caption = "Υπολογίζονται τα κενά στοιχεία του προγράμματος..."
    DoEvents
    response = CompletePatientProgram(mPatientId, requestArguments)'''
    _patch_code_module(
        code_module,
        auto_validation_old,
        auto_validation_new,
        description="patient-program auto conflict validation",
    )

    save_decl_old = '''    Dim writesStarted As Boolean
    Dim errorText As String

    SaveCurrentDraft'''
    save_decl_new = '''    Dim writesStarted As Boolean
    Dim errorText As String
    Dim conflictText As String

    SaveCurrentDraft'''
    _patch_code_module(
        code_module,
        save_decl_old,
        save_decl_new,
        description="patient-program save conflict declaration",
    )

    save_validation_old = '''    If Not DraftIsComplete() Then
        Err.Raise vbObjectError + 2930, "SavePatientProgramAtomic", _
            "Το πρόγραμμα δεν είναι ακόμη πλήρες. Χρησιμοποίησε Αυτόματη συμπλήρωση ή συμπλήρωσε τα κενά πεδία."
    End If

    Set ws = ThisWorkbook.Worksheets("PATIENT_PLANNER")'''
    save_validation_new = '''    If Not DraftIsComplete() Then
        Err.Raise vbObjectError + 2930, "SavePatientProgramAtomic", _
            "Το πρόγραμμα δεν είναι ακόμη πλήρες. Χρησιμοποίησε Αυτόματη συμπλήρωση ή συμπλήρωσε τα κενά πεδία."
    End If

    conflictText = CurrentConflictText()
    If Len(conflictText) > 0 Then
        If Not ConfirmConflictOverride(conflictText) Then
            lblStatus.Caption = "Η αποθήκευση ακυρώθηκε για να διορθωθεί η σύγκρουση."
            Exit Sub
        End If
    End If

    Set ws = ThisWorkbook.Worksheets("PATIENT_PLANNER")'''
    _patch_code_module(
        code_module,
        save_validation_old,
        save_validation_new,
        description="patient-program save conflict validation",
    )


def _patch_patient_edit_refresh(code_module) -> None:
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    raw_text = code_module.Lines(1, count) if count else ""
    text = _normalize_newlines(raw_text)

    procedure_marker = "Private Sub cmdSave_Click()"
    procedure_start = text.find(procedure_marker)
    if procedure_start < 0:
        raise RuntimeError("MASTER render patch marker missing: patient-edit cmdSave_Click")

    procedure_end = text.find("\nEnd Sub", procedure_start)
    if procedure_end < 0:
        raise RuntimeError("MASTER render patch marker missing: patient-edit cmdSave_Click end")

    block = text[procedure_start:procedure_end]
    refresh_line = "    RefreshMasterPatientProgram patientId"
    if refresh_line in block:
        return

    success_marker = '    MsgBox "Οι αλλαγές αποθηκεύτηκαν επιτυχώς."'
    marker_index = block.find(success_marker)
    if marker_index < 0:
        raise RuntimeError("MASTER render patch marker missing: patient-edit success message")

    patched_block = block[:marker_index] + refresh_line + "\n\n" + block[marker_index:]
    patched_text = text[:procedure_start] + patched_block + text[procedure_end:]
    _write_code_module(code_module, patched_text)


def install_patient_program_master_render(vbproject) -> None:
    """Install MASTER renderer plus patient-program conflict validation hooks."""

    module = _get_component(vbproject, MASTER_RENDER_MODULE_NAME)
    if module is None:
        module = vbproject.VBComponents.Add(1)
        module.Name = MASTER_RENDER_MODULE_NAME
    _replace_component_code(module, MASTER_RENDER_MODULE_CODE)

    _install_conflict_form(vbproject)

    runtime = vbproject.VBComponents(PATIENT_PROGRAM_RUNTIME_NAME)
    _patch_runtime_validation(runtime.CodeModule)

    program_form = vbproject.VBComponents(PATIENT_PROGRAM_FORM_NAME)
    _patch_program_conflict_validation(program_form.CodeModule)
    _patch_code_module(
        program_form.CodeModule,
        "    ThisWorkbook.Save\n    Application.EnableEvents = previousEvents",
        "    RefreshMasterPatientProgram mPatientId\n    ThisWorkbook.Save\n    Application.EnableEvents = previousEvents",
        description="patient-program save refresh",
    )

    edit_form = vbproject.VBComponents(PATIENT_EDIT_FORM_NAME)
    _patch_patient_edit_refresh(edit_form.CodeModule)
