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
    new_doctor_write = '    patients.Cells(targetRow, doctorCol).Value = Trim$(responsibleDoctor)'
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
    return code.replace(old_name, new_name, 1)


def _vertical_form_code() -> str:
    code = _diagnostics._patched_form_code()

    validation_marker = "    If Not ValidateForm() Then Exit Sub\n\n    patientType = Trim$(cboPatientType.Value)"
    validation = '''    If Not ValidateForm() Then Exit Sub

    If Len(Trim$(txtResponsibleDoctor.Text)) = 0 Then
        MsgBox "Ο υπεύθυνος γιατρός είναι υποχρεωτικός.", vbExclamation, "Νέος ασθενής"
        txtResponsibleDoctor.SetFocus
        Exit Sub
    End If

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
    """Add the doctor controls to frmNewPatient without changing legacy sources.

    The form is created by the registration-menu build stage. This augmentation
    is build-time only; normal registration remains fully in-session.
    """

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

        code_module = component.CodeModule
        form_code = code_module.Lines(1, code_module.CountOfLines)
        marker = "    LoadSettingsValues\n    ApplyPatientTypeRules"
        injected = '''    Me.Height = 610
    StyleLabel lblResponsibleDoctor, "Υπεύθυνος γιατρός *", 370
    StyleTextBox txtResponsibleDoctor, 366
    lblInfo.Top = 414
    lblRequired.Top = 458
    cmdCancel.Top = 496
    cmdSave.Top = 496

    LoadSettingsValues
    ApplyPatientTypeRules'''
        if "StyleLabel lblResponsibleDoctor" not in form_code:
            if marker not in form_code:
                raise PatientRegistrationBridgeVbaError(
                    "Responsible-doctor form initialization marker was not found"
                )
            form_code = form_code.replace(marker, injected, 1)
            code_module.DeleteLines(1, code_module.CountOfLines)
            code_module.AddFromString(form_code)

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
