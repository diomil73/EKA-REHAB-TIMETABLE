from __future__ import annotations

"""Patient-centric registration slice layered on the validated V36 fast path.

This keeps the successful in-session registration architecture, adds the
responsible doctor to the same patient workflow, keeps MASTER projection
blank-safe, and exposes an in-session patient edit backend.
"""

from pathlib import Path
import sys

from . import patient_registration_bridge_diagnostics as _diagnostics
from . import patient_registration_bridge_vba as _base

PatientRegistrationBridgeVbaError = _base.PatientRegistrationBridgeVbaError


def _vertical_module_code() -> str:
    code = _diagnostics._patched_module_code()

    old_signature = '''    ByVal roomValue As String, _
    ByVal infectious As Boolean, _
    ByVal statusValue As String _
) As String'''
    new_signature = '''    ByVal roomValue As String, _
    ByVal infectious As Boolean, _
    ByVal responsibleDoctor As String, _
    ByVal statusValue As String _
) As String'''
    if old_signature not in code:
        raise RuntimeError("Responsible-doctor function signature marker missing")
    code = code.replace(old_signature, new_signature, 1)

    old_doctor_write = '    patients.Cells(targetRow, doctorCol).Value = ""'
    new_doctor_write = '''    If IsOutpatient(patientType) Then
        patients.Cells(targetRow, doctorCol).Value = ""
    Else
        patients.Cells(targetRow, doctorCol).Value = Trim$(responsibleDoctor)
    End If'''
    if old_doctor_write not in code:
        raise RuntimeError("Responsible-doctor PATIENTS write marker missing")
    code = code.replace(old_doctor_write, new_doctor_write, 1)

    edit_marker = "Private Function IsOutpatient(ByVal patientType As String) As Boolean"
    edit_backend = r'''Public Sub UpdatePatientInWorkbook( _
    ByVal patientId As String, _
    ByVal hospitalMRN As String, _
    ByVal displayName As String, _
    ByVal roomValue As String, _
    ByVal infectious As Boolean, _
    ByVal responsibleDoctor As String, _
    ByVal statusValue As String _
)
    Dim patients As Worksheet
    Dim planner As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim typeCol As Long
    Dim mrnCol As Long
    Dim doctorCol As Long
    Dim patientType As String
    Dim candidateMrn As String
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo UpdateError

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set patients = ThisWorkbook.Worksheets("PATIENTS")
    Set planner = ThisWorkbook.Worksheets("PATIENT_PLANNER")

    typeCol = FindOrCreateHeader(patients, Array("PatientType", "ΤύποςΑσθενή", "Τύπος Ασθενή"), "PatientType")
    mrnCol = FindOrCreateHeader(patients, Array("HospitalMRN", "ΑΜ Νοσοκομείου", "ΑΜΝοσοκομείου"), "HospitalMRN")
    doctorCol = FindOrCreateHeader(patients, Array("ResponsibleDoctor", "ΥπεύθυνοςΙατρός", "Υπεύθυνος Ιατρός", "ΥπεύθυνοςΓιατρός", "Υπεύθυνος Γιατρός"), "ResponsibleDoctor")

    targetRow = FindPatientRowById(patients, patientId)
    If targetRow = 0 Then Err.Raise vbObjectError + 2460, , "Δεν βρέθηκε ο επιλεγμένος ασθενής."

    patientType = Trim$(CStr(patients.Cells(targetRow, typeCol).Value))
    If Len(Trim$(displayName)) = 0 Then Err.Raise vbObjectError + 2461, , "Το ονοματεπώνυμο είναι υποχρεωτικό."
    If Len(Trim$(statusValue)) = 0 Then Err.Raise vbObjectError + 2462, , "Η κατάσταση είναι υποχρεωτική."
    If Not IsOutpatient(patientType) And Len(Trim$(roomValue)) = 0 Then Err.Raise vbObjectError + 2463, , "Ο θάλαμος είναι υποχρεωτικός για εσωτερικό ασθενή."
    If Not IsOutpatient(patientType) And Len(Trim$(responsibleDoctor)) = 0 Then Err.Raise vbObjectError + 2464, , "Ο υπεύθυνος γιατρός είναι υποχρεωτικός για εσωτερικό ασθενή."

    candidateMrn = Trim$(hospitalMRN)
    If Len(candidateMrn) > 0 Then
        For rowIndex = 2 To LastNonBlankRow(patients, mrnCol)
            If rowIndex <> targetRow Then
                If StrComp(Trim$(CStr(patients.Cells(rowIndex, mrnCol).Value)), candidateMrn, vbTextCompare) = 0 Then
                    Err.Raise vbObjectError + 2465, , "Ο ΑΜ Νοσοκομείου υπάρχει ήδη σε άλλον ασθενή."
                End If
            End If
        Next rowIndex
    End If

    patients.Cells(targetRow, 3).Value = Trim$(displayName)
    patients.Cells(targetRow, 5).Value = Trim$(statusValue)
    patients.Cells(targetRow, mrnCol).Value = candidateMrn

    If IsOutpatient(patientType) Then
        patients.Cells(targetRow, 2).Value = ""
        patients.Cells(targetRow, 4).Value = ""
        patients.Cells(targetRow, doctorCol).Value = ""
    Else
        patients.Cells(targetRow, 2).Value = Trim$(roomValue)
        If infectious Then
            patients.Cells(targetRow, 4).Value = "Ν"
        Else
            patients.Cells(targetRow, 4).Value = ""
        End If
        patients.Cells(targetRow, doctorCol).Value = Trim$(responsibleDoctor)
    End If

    RefreshPlannerRow planner, targetRow, IsOutpatient(patientType)
    RebuildMasterFast patients, typeCol, doctorCol
    planner.Range("A" & CStr(targetRow) & ":E" & CStr(targetRow)).Calculate
    ThisWorkbook.Save

UpdateExit:
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Sub

UpdateError:
    Dim errorText As String
    errorText = Err.Description
    On Error Resume Next
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 2469, "UpdatePatientInWorkbook", errorText
End Sub

Private Function FindPatientRowById(ByVal ws As Worksheet, ByVal patientId As String) As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = LastNonBlankRow(ws, 1)
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 1).Value)), Trim$(patientId), vbTextCompare) = 0 Then
            FindPatientRowById = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

'''
    if edit_marker not in code:
        raise RuntimeError("Patient edit backend injection marker missing")
    code = code.replace(edit_marker, edit_backend + edit_marker, 1)

    # Direct Excel links turn blank source cells into visible zeroes. Keep the
    # links dynamic, but explicitly return an empty string for an empty source.
    for master_col, planner_col in (
        (1, "D"),
        (2, "B"),
        (4, "F"),
        (5, "I"),
        (6, "N"),
        (7, "L"),
        (8, "P"),
        (9, "R"),
        (10, "T"),
    ):
        old = (
            f'        ws.Cells(targetRow, {master_col}).Formula = '
            f'"=PATIENT_PLANNER!{planner_col}" & plannerRow'
        )
        new = (
            f'        ws.Cells(targetRow, {master_col}).Formula = '
            f'"=IF(PATIENT_PLANNER!{planner_col}" & plannerRow & '
            f'"="""","""",PATIENT_PLANNER!{planner_col}" & plannerRow & ")"'
        )
        if old not in code:
            raise RuntimeError(f"Blank-safe MASTER formula marker missing for column {master_col}")
        code = code.replace(old, new, 1)

    old_status = (
        '        ws.Cells(targetRow, statusCol).Formula = '
        '"=PATIENT_PLANNER!E" & plannerRow'
    )
    new_status = (
        '        ws.Cells(targetRow, statusCol).Formula = '
        '"=IF(PATIENT_PLANNER!E" & plannerRow & '
        '"="""","""",PATIENT_PLANNER!E" & plannerRow & ")"'
    )
    if old_status not in code:
        raise RuntimeError("Blank-safe MASTER status formula marker missing")
    code = code.replace(old_status, new_status, 1)

    old_name = '            ws.Cells(targetRow, 3).Formula = "=PATIENT_PLANNER!C" & plannerRow'
    new_name = (
        '            ws.Cells(targetRow, 3).Formula = '
        '"=IF(PATIENT_PLANNER!C" & plannerRow & '
        '"="""","""",PATIENT_PLANNER!C" & plannerRow & ")"'
    )
    if old_name not in code:
        raise RuntimeError("Blank-safe MASTER patient-name formula marker missing")
    code = code.replace(old_name, new_name, 1)

    old_doctor_block = '''        If Len(doctor) > 0 Then
            ws.Cells(targetRow, 3).Value = displayName & vbLf & "✚ " & doctor
        Else
            ws.Cells(targetRow, 3).Formula = "=IF(PATIENT_PLANNER!C" & plannerRow & "="""","""",PATIENT_PLANNER!C" & plannerRow & ")"
        End If

        ApplyMasterPatientStyle ws, targetRow, IsTruthy(patients.Cells(plannerRow, 4).Value)'''
    new_doctor_block = '''        If Len(doctor) > 0 Then
            WritePatientDoctorCell ws.Cells(targetRow, 3), displayName, doctor
        Else
            ws.Cells(targetRow, 3).Formula = "=IF(PATIENT_PLANNER!C" & plannerRow & "="""","""",PATIENT_PLANNER!C" & plannerRow & ")"
        End If

        ApplyMasterPatientStyle ws, targetRow, IsTruthy(patients.Cells(plannerRow, 4).Value)
        If Len(doctor) > 0 Then ApplyDoctorLineStyle ws.Cells(targetRow, 3), displayName, doctor'''
    if old_doctor_block not in code:
        raise RuntimeError("Responsible-doctor MASTER display marker missing")
    code = code.replace(old_doctor_block, new_doctor_block, 1)

    style_marker = "Private Sub RebuildMasterFast(ByVal patients As Worksheet, ByVal typeCol As Long, ByVal doctorCol As Long)"
    doctor_style = '''Private Function DoctorSeparator() As String
    Dim i As Long
    Dim text As String

    For i = 1 To 24
        text = text & ChrW$(9472)
    Next i
    DoctorSeparator = text
End Function

Private Sub WritePatientDoctorCell(ByVal targetCell As Range, ByVal displayName As String, ByVal doctor As String)
    Dim separator As String

    separator = DoctorSeparator()
    targetCell.Value = displayName & vbLf & separator & vbLf & doctor
End Sub

Private Sub ApplyDoctorLineStyle(ByVal targetCell As Range, ByVal displayName As String, ByVal doctor As String)
    Dim separator As String
    Dim separatorStart As Long
    Dim doctorStart As Long

    separator = DoctorSeparator()
    separatorStart = Len(displayName) + 2
    doctorStart = Len(displayName) + Len(separator) + 3

    With targetCell
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With

    With targetCell.Characters(Start:=1, Length:=Len(displayName)).Font
        .Name = "Segoe UI"
        .Size = 10
        .Bold = True
        .Italic = False
    End With

    With targetCell.Characters(Start:=separatorStart, Length:=Len(separator)).Font
        .Name = "Segoe UI"
        .Size = 6
        .Bold = False
        .Italic = False
        .Color = RGB(180, 185, 195)
    End With

    With targetCell.Characters(Start:=doctorStart, Length:=Len(doctor)).Font
        .Name = "Segoe UI"
        .Size = 9
        .Bold = False
        .Italic = True
        .Color = RGB(92, 64, 120)
    End With
End Sub

'''
    if style_marker not in code:
        raise RuntimeError("MASTER rebuild marker missing for doctor style injection")
    code = code.replace(style_marker, doctor_style + style_marker, 1)

    code = code.replace(
        '    ws.Rows(rowIndex).RowHeight = 64',
        '    ws.Rows(rowIndex).RowHeight = 72',
        1,
    )
    return code


def _vertical_form_code() -> str:
    code = _diagnostics._patched_form_code()

    validation_marker = "    If Not ValidateForm() Then Exit Sub\n\n    patientType = Trim$(cboPatientType.Value)"
    validation = '''    If Not ValidateForm() Then Exit Sub

    patientType = Trim$(cboPatientType.Value)'''
    if validation_marker not in code:
        raise RuntimeError("Responsible-doctor validation marker missing")
    code = code.replace(validation_marker, validation, 1)

    old = '''        roomValue, _
        CBool(chkInfectious.Value), _
        Trim$(cboStatus.Value) _'''
    new = '''        roomValue, _
        CBool(chkInfectious.Value), _
        Trim$(txtResponsibleDoctor.Text), _
        Trim$(cboStatus.Value) _'''
    if old not in code:
        raise RuntimeError("Responsible-doctor form-call marker missing")
    return code.replace(old, new, 1)


def _augment_patient_form(workbook_path: Path) -> None:
    """Keep compatibility with previews built from older registration-menu stages."""

    if sys.platform != "win32":
        raise PatientRegistrationBridgeVbaError(
            "Patient form augmentation requires Windows with Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientRegistrationBridgeVbaError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)
        component = workbook.VBProject.VBComponents(_base.PATIENT_FORM_NAME)
        designer = component.Designer

        try:
            designer.Controls("lblResponsibleDoctor")
        except Exception:
            label = designer.Controls.Add("Forms.Label.1", "lblResponsibleDoctor", True)
            label.Caption = "Υπεύθυνος γιατρός *"

        try:
            designer.Controls("txtResponsibleDoctor")
        except Exception:
            designer.Controls.Add("Forms.TextBox.1", "txtResponsibleDoctor", True)

        workbook.Save()
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Responsible-doctor form augmentation failed: {exc}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def create_patient_registration_bridge_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
):
    original_module = _base.BRIDGE_MODULE_CODE
    original_form = _base.FORM_BRIDGE_CODE
    try:
        _base.BRIDGE_MODULE_CODE = _vertical_module_code()
        _base.FORM_BRIDGE_CODE = _vertical_form_code()
        report = _base.create_patient_registration_bridge_preview(
            source_path,
            output_path,
            overwrite=overwrite,
        )
        _augment_patient_form(Path(output_path))
        return report
    finally:
        _base.BRIDGE_MODULE_CODE = original_module
        _base.FORM_BRIDGE_CODE = original_form
