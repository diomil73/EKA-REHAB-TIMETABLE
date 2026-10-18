from __future__ import annotations

from pathlib import Path
import sys

EDIT_MODULE_NAME = "modOutpatientScheduleEdit"
EDIT_MACRO_NAME = "EditOutpatientSchedule"
FORM_NAME = "frmOutpatientSchedule"
FORM_MODULE_NAME = "modOutpatientSchedule"

EDIT_MODULE_CODE = r'''Option Explicit

Public OutpatientEditTargetBaseEntryId As String

Public Sub EditOutpatientSchedule()
    Dim ws As Worksheet
    Dim defaultRow As Long
    Dim rowText As String
    Dim rowIndex As Long
    Dim patientId As String
    Dim patientName As String
    Dim treatment As String
    Dim dayPattern As String
    Dim therapistId As String
    Dim timeText As String

    On Error GoTo EditError

    Set ws = ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")
    defaultRow = 2
    If ActiveSheet.Name = "OUTPATIENT_SCHEDULE" Then
        If ActiveCell.Row >= 2 Then defaultRow = ActiveCell.Row
    End If

    rowText = InputBox( _
        "Δώστε τη γραμμή του OUTPATIENT_SCHEDULE που θέλετε να επεξεργαστείτε.", _
        "Επεξεργασία προγράμματος εξωτερικού ασθενή", _
        CStr(defaultRow) _
    )
    If Len(Trim$(rowText)) = 0 Then Exit Sub
    If Not IsNumeric(rowText) Then
        MsgBox "Η γραμμή πρέπει να είναι αριθμός.", vbExclamation
        Exit Sub
    End If

    rowIndex = CLng(rowText)
    If rowIndex < 2 Or rowIndex > ws.Cells(ws.Rows.Count, 1).End(xlUp).Row Then
        MsgBox "Δεν υπάρχει εγγραφή σε αυτή τη γραμμή.", vbExclamation
        Exit Sub
    End If

    patientId = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
    patientName = Trim$(CStr(ws.Cells(rowIndex, 2).Value))
    treatment = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
    dayPattern = Trim$(CStr(ws.Cells(rowIndex, 5).Value))
    therapistId = Trim$(CStr(ws.Cells(rowIndex, 6).Value))

    If Len(patientId) = 0 Or Len(patientName) = 0 Or Len(treatment) = 0 Or Len(dayPattern) = 0 Then
        MsgBox "Η επιλεγμένη γραμμή δεν είναι έγκυρη εγγραφή εξωτερικού προγράμματος.", vbExclamation
        Exit Sub
    End If

    If IsDate(ws.Cells(rowIndex, 4).Value) Then
        timeText = Format$(ws.Cells(rowIndex, 4).Value, "hh:mm")
    Else
        timeText = Trim$(CStr(ws.Cells(rowIndex, 4).Text))
    End If
    If Len(timeText) = 0 Then
        MsgBox "Η επιλεγμένη εγγραφή δεν έχει έγκυρη ώρα.", vbExclamation
        Exit Sub
    End If

    Load frmOutpatientSchedule
    OutpatientEditTargetBaseEntryId = "outpatient:" & CStr(rowIndex)

    With frmOutpatientSchedule
        .cboPatient.Value = patientId & " | " & patientName
        .cboTreatment.Value = treatment
        .cboTime.Value = timeText
        .cboDays.Value = dayPattern
        .cboTherapist.Value = therapistId
        .Caption = "Επεξεργασία προγράμματος εξωτερικού ασθενή"
        .lblTitle.Caption = "Επεξεργασία επαναλαμβανόμενου προγράμματος"
        .lblInfo.Caption = "Επεξεργάζεστε την εγγραφή " & OutpatientEditTargetBaseEntryId & ". Οι αλλαγές θα ενημερώσουν την ίδια γραμμή μετά τον έλεγχο."
        .cmdSave.Caption = "Έλεγχος αλλαγών και preview"
        .Show
    End With
    Exit Sub

EditError:
    MsgBox "Δεν ήταν δυνατό να ανοίξει η εγγραφή για επεξεργασία: " & Err.Description, vbCritical
End Sub
'''

_INITIALIZE_MARKER = '''Private Sub UserForm_Initialize()\n    Me.Caption = "Πρόγραμμα εξωτερικού ασθενή"'''
_INITIALIZE_REPLACEMENT = '''Private Sub UserForm_Initialize()\n    OutpatientEditTargetBaseEntryId = ""\n    Me.Caption = "Πρόγραμμα εξωτερικού ασθενή"'''

_REQUEST_MARKER = '''        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & _\n        "}}"'''
_REQUEST_REPLACEMENT = '''        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & "," & _\n        q & "target_base_entry_id" & q & ":" & q & JsonEscape(OutpatientEditTargetBaseEntryId) & q & _\n        "}}"'''

_SUGGESTION_MARKER = '''        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _\n        q & "limit" & q & ":5" & _'''
_SUGGESTION_REPLACEMENT = '''        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _\n        q & "target_base_entry_id" & q & ":" & q & JsonEscape(OutpatientEditTargetBaseEntryId) & q & "," & _\n        q & "limit" & q & ":5" & _'''


def patch_outpatient_schedule_form_code(code: str) -> str:
    """Add edit-target state to the existing outpatient schedule form code.

    VBIDE commonly returns CRLF text even when the source template uses LF. Normalize
    first so structural markers are stable in both pure-Python tests and native Excel.
    """

    patched = code.replace("\r\n", "\n").replace("\r", "\n")
    if 'OutpatientEditTargetBaseEntryId = ""' not in patched:
        if _INITIALIZE_MARKER not in patched:
            raise RuntimeError("Outpatient schedule form initialize marker missing")
        patched = patched.replace(_INITIALIZE_MARKER, _INITIALIZE_REPLACEMENT, 1)

    if patched.count('q & "target_base_entry_id" & q') < 1:
        if _REQUEST_MARKER not in patched:
            raise RuntimeError("Outpatient schedule request JSON marker missing")
        patched = patched.replace(_REQUEST_MARKER, _REQUEST_REPLACEMENT, 1)

    if patched.count('q & "target_base_entry_id" & q') < 2:
        if _SUGGESTION_MARKER not in patched:
            raise RuntimeError("Outpatient suggestion request JSON marker missing")
        patched = patched.replace(_SUGGESTION_MARKER, _SUGGESTION_REPLACEMENT, 1)

    return patched


def _remove_component_if_present(vbproject, component_name: str) -> None:
    try:
        component = vbproject.VBComponents(component_name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _component_code(component) -> str:
    module = component.CodeModule
    count = int(getattr(module, "CountOfLines", 0) or 0)
    if count <= 0:
        return ""
    return str(module.Lines(1, count))


def install_outpatient_schedule_edit_support(vbproject) -> None:
    """Install edit/preload support while reusing the existing form and writer."""

    form_component = vbproject.VBComponents(FORM_NAME)
    current_form_code = _component_code(form_component)
    patched_form_code = patch_outpatient_schedule_form_code(current_form_code)

    if patched_form_code != current_form_code:
        form_module = form_component.CodeModule
        if form_module.CountOfLines:
            form_module.DeleteLines(1, form_module.CountOfLines)
        form_module.AddFromString(patched_form_code)

    _remove_component_if_present(vbproject, EDIT_MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = EDIT_MODULE_NAME
    module.CodeModule.AddFromString(EDIT_MODULE_CODE)

    installed_form_code = _component_code(form_component)
    installed_edit_code = _component_code(module)
    if installed_form_code.count('q & "target_base_entry_id" & q') < 2:
        raise RuntimeError("Outpatient edit target was not wired into save and suggestion JSON")
    if "OutpatientEditTargetBaseEntryId" not in installed_form_code:
        raise RuntimeError("Outpatient edit target state was not wired into form")
    if "Public Sub EditOutpatientSchedule()" not in installed_edit_code:
        raise RuntimeError("Outpatient edit macro was not installed")


def install_outpatient_schedule_edit_support_in_workbook(workbook_path: str | Path) -> None:
    """Patch an already-built outpatient schedule form preview in place."""

    path = Path(workbook_path).resolve()
    if sys.platform != "win32":
        raise RuntimeError("Outpatient schedule edit support requires Windows Excel")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        install_outpatient_schedule_edit_support(workbook.VBProject)
        workbook.Save()
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
