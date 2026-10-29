from __future__ import annotations

"""Targeted V35/V36 diagnostics for the in-session patient-registration path.

This module deliberately wraps the proven bridge installer instead of duplicating
its COM/VBProject wiring.  The generated VBA stays in-session, but failures now
report the exact stage and the logical PATIENTS row search ignores physically
extended blank/table rows.
"""

from pathlib import Path

from . import patient_registration_bridge_vba as _base

PatientRegistrationBridgeVbaError = _base.PatientRegistrationBridgeVbaError


def _patched_module_code() -> str:
    code = _base.BRIDGE_MODULE_CODE

    code = code.replace(
        "    Dim previousCalculation As XlCalculation\n\n    On Error GoTo RegistrationError",
        "    Dim previousCalculation As XlCalculation\n"
        "    Dim stage As String\n\n"
        "    stage = \"initialization\"\n"
        "    On Error GoTo RegistrationError",
    )

    replacements = (
        (
            '    Set patients = ThisWorkbook.Worksheets("PATIENTS")',
            '    stage = "open worksheets"\n    Set patients = ThisWorkbook.Worksheets("PATIENTS")',
        ),
        (
            '    typeCol = FindOrCreateHeader(patients, Array(',
            '    stage = "resolve PATIENTS schema"\n    typeCol = FindOrCreateHeader(patients, Array(',
        ),
        (
            '    ValidateFastRegistration patients, mrnCol, patientType, hospitalMRN, displayName, roomValue, statusValue',
            '    stage = "validate patient"\n    ValidateFastRegistration patients, mrnCol, patientType, hospitalMRN, displayName, roomValue, statusValue',
        ),
        (
            '    patientId = NextPatientId(patients)\n    targetRow = NextPatientTargetRow(patients)',
            '    stage = "allocate patient id and row"\n    patientId = NextPatientId(patients)\n    targetRow = NextPatientTargetRow(patients)',
        ),
        (
            '    EnsurePatientTableRow patients, targetRow\n    EnsureOutpatientScheduleSheet',
            '    stage = "prepare registry rows"\n    EnsurePatientTableRow patients, targetRow\n    EnsureOutpatientScheduleSheet',
        ),
        (
            '    patients.Cells(targetRow, 1).Value = patientId',
            '    stage = "write PATIENTS row " & CStr(targetRow)\n    patients.Cells(targetRow, 1).Value = patientId',
        ),
        (
            '    RefreshPlannerRow planner, targetRow, IsOutpatient(patientType)',
            '    stage = "refresh PATIENT_PLANNER row " & CStr(targetRow)\n    RefreshPlannerRow planner, targetRow, IsOutpatient(patientType)',
        ),
        (
            '    If Not IsOutpatient(patientType) Then\n        RebuildMasterFast patients, typeCol, doctorCol\n    End If',
            '    If Not IsOutpatient(patientType) Then\n        stage = "rebuild MASTER"\n        RebuildMasterFast patients, typeCol, doctorCol\n    End If',
        ),
        (
            '    planner.Range("A" & CStr(targetRow) & ":E" & CStr(targetRow)).Calculate\n    ThisWorkbook.Save',
            '    stage = "calculate registered row"\n    planner.Range("A" & CStr(targetRow) & ":E" & CStr(targetRow)).Calculate\n    stage = "save workbook"\n    ThisWorkbook.Save',
        ),
        (
            '    errorText = Err.Description',
            '    errorText = "Stage: " & stage & vbCrLf & Err.Description',
        ),
    )
    for old, new in replacements:
        if old not in code:
            raise RuntimeError(f"Patient registration diagnostic patch marker missing: {old[:80]}")
        code = code.replace(old, new, 1)

    old_target = '''Private Function NextPatientTargetRow(ByVal ws As Worksheet) As Long
    Dim lastId As Long
    Dim lastName As Long
    lastId = LastNonBlankRow(ws, 1)
    lastName = LastNonBlankRow(ws, 3)
    NextPatientTargetRow = Application.Max(2, Application.Max(lastId, lastName) + 1)
End Function'''
    new_target = '''Private Function LastLogicalPatientRow(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    Dim physicalLast As Long
    Dim rowIndex As Long
    Dim valueText As String

    physicalLast = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    For rowIndex = physicalLast To 2 Step -1
        valueText = Trim$(CStr(ws.Cells(rowIndex, columnIndex).Value2))
        If Len(valueText) > 0 Then
            LastLogicalPatientRow = rowIndex
            Exit Function
        End If
    Next rowIndex
    LastLogicalPatientRow = 1
End Function

Private Function NextPatientTargetRow(ByVal ws As Worksheet) As Long
    Dim lastId As Long
    Dim lastName As Long
    lastId = LastLogicalPatientRow(ws, 1)
    lastName = LastLogicalPatientRow(ws, 3)
    NextPatientTargetRow = Application.Max(2, Application.Max(lastId, lastName) + 1)
End Function'''
    if old_target not in code:
        raise RuntimeError("NextPatientTargetRow diagnostic patch marker missing")
    return code.replace(old_target, new_target, 1)


def _patched_form_code() -> str:
    code = _base.FORM_BRIDGE_CODE
    old = '''    MsgBox "Ο ασθενής καταχωρήθηκε επιτυχώς." & vbCrLf & _
           "Patient ID: " & patientId, _
           vbInformation, "Νέος ασθενής"
    Unload Me
    Exit Sub'''
    new = '''    On Error Resume Next
    Unload frmRegistrationMenu
    On Error GoTo RegistrationError
    Unload Me
    GoToMaster
    MsgBox "Ο ασθενής καταχωρήθηκε επιτυχώς." & vbCrLf & _
           "Patient ID: " & patientId, _
           vbInformation, "Νέος ασθενής"
    Exit Sub'''
    if old not in code:
        raise RuntimeError("Patient registration success-navigation patch marker missing")
    return code.replace(old, new, 1)


def create_patient_registration_bridge_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
):
    original_module = _base.BRIDGE_MODULE_CODE
    original_form = _base.FORM_BRIDGE_CODE
    try:
        _base.BRIDGE_MODULE_CODE = _patched_module_code()
        _base.FORM_BRIDGE_CODE = _patched_form_code()
        return _base.create_patient_registration_bridge_preview(
            source_path,
            output_path,
            overwrite=overwrite,
        )
    finally:
        _base.BRIDGE_MODULE_CODE = original_module
        _base.FORM_BRIDGE_CODE = original_form
