from __future__ import annotations

"""Patient-centric registration slice layered on the validated V36 fast path.

This keeps the successful in-session registration architecture, adds the
responsible doctor to the same patient workflow, and makes MASTER projection
links blank-safe so empty PATIENT_PLANNER cells do not render as zeroes.
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
            ws.Cells(targetRow, 3).Value = displayName & vbLf & doctor
        Else
            ws.Cells(targetRow, 3).Formula = "=IF(PATIENT_PLANNER!C" & plannerRow & "="""","""",PATIENT_PLANNER!C" & plannerRow & ")"
        End If

        ApplyMasterPatientStyle ws, targetRow, IsTruthy(patients.Cells(plannerRow, 4).Value)
        If Len(doctor) > 0 Then ApplyDoctorLineStyle ws.Cells(targetRow, 3), displayName, doctor'''
    if old_doctor_block not in code:
        raise RuntimeError("Responsible-doctor MASTER display marker missing")
    code = code.replace(old_doctor_block, new_doctor_block, 1)

    style_marker = "Private Sub RebuildMasterFast(ByVal patients As Worksheet, ByVal typeCol As Long, ByVal doctorCol As Long)"
    doctor_style = '''Private Sub ApplyDoctorLineStyle(ByVal targetCell As Range, ByVal displayName As String, ByVal doctor As String)
    Dim doctorStart As Long

    doctorStart = Len(displayName) + 2
    With targetCell
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With

    With targetCell.Characters(Start:=1, Length:=Len(displayName)).Font
        .Name = "Calibri"
        .Size = 10.5
        .Bold = True
        .Italic = False
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
    return code.replace(style_marker, doctor_style + style_marker, 1)


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
