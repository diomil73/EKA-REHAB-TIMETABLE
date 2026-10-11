from __future__ import annotations


MASTER_RENDER_MODULE_NAME = "PatientProgramMasterRender"
PATIENT_PROGRAM_FORM_NAME = "frmPatientScheduling"
PATIENT_EDIT_FORM_NAME = "frmEditPatient"


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
    Dim parts() As String
    parts = Split(Replace(valueText, vbCrLf, vbLf), vbLf)
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
    """Install MASTER patient-program renderer and wire save/edit refresh calls."""

    module = _get_component(vbproject, MASTER_RENDER_MODULE_NAME)
    if module is None:
        module = vbproject.VBComponents.Add(1)
        module.Name = MASTER_RENDER_MODULE_NAME
    _replace_component_code(module, MASTER_RENDER_MODULE_CODE)

    program_form = vbproject.VBComponents(PATIENT_PROGRAM_FORM_NAME)
    _patch_code_module(
        program_form.CodeModule,
        "    ThisWorkbook.Save\n    Application.EnableEvents = previousEvents",
        "    RefreshMasterPatientProgram mPatientId\n    ThisWorkbook.Save\n    Application.EnableEvents = previousEvents",
        description="patient-program save refresh",
    )

    edit_form = vbproject.VBComponents(PATIENT_EDIT_FORM_NAME)
    _patch_patient_edit_refresh(edit_form.CodeModule)
