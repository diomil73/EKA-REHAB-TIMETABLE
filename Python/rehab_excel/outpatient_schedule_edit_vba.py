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
    ShowOutpatientScheduleManager
End Sub

Public Sub EditOutpatientScheduleByRow(ByVal rowIndex As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim patientId As String
    Dim patientName As String
    Dim treatment As String
    Dim dayPattern As String
    Dim therapistId As String
    Dim timeText As String

    On Error GoTo EditError

    Set ws = ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If rowIndex < 2 Or rowIndex > lastRow Then
        MsgBox "Η επιλεγμένη εγγραφή δεν είναι πλέον διαθέσιμη.", vbExclamation, _
               "Επεξεργασία προγράμματος εξωτερικού ασθενή"
        Exit Sub
    End If

    patientId = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
    patientName = Trim$(CStr(ws.Cells(rowIndex, 2).Value))
    treatment = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
    dayPattern = Trim$(CStr(ws.Cells(rowIndex, 5).Value))
    therapistId = Trim$(CStr(ws.Cells(rowIndex, 6).Value))

    If Len(patientId) = 0 Or Len(patientName) = 0 Or Len(treatment) = 0 Or Len(dayPattern) = 0 Then
        MsgBox "Η επιλεγμένη εγγραφή δεν είναι έγκυρο πρόγραμμα εξωτερικού ασθενή.", _
               vbExclamation, "Επεξεργασία προγράμματος εξωτερικού ασθενή"
        Exit Sub
    End If

    If IsDate(ws.Cells(rowIndex, 4).Value) Then
        timeText = Format$(ws.Cells(rowIndex, 4).Value, "hh:mm")
    Else
        timeText = Trim$(CStr(ws.Cells(rowIndex, 4).Text))
    End If
    If Len(timeText) = 0 Then
        MsgBox "Η επιλεγμένη εγγραφή δεν έχει έγκυρη ώρα.", vbExclamation, _
               "Επεξεργασία προγράμματος εξωτερικού ασθενή"
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
        .lblInfo.Caption = "Επεξεργαστείτε τα στοιχεία και αποθηκεύστε τις αλλαγές."
        .cmdSave.Caption = "Αποθήκευση αλλαγών"
        .Show
    End With
    Exit Sub

EditError:
    MsgBox "Δεν ήταν δυνατό να ανοίξει η εγγραφή για επεξεργασία: " & Err.Description, _
           vbCritical, "Επεξεργασία προγράμματος εξωτερικού ασθενή"
End Sub
'''

_REQUEST_PROCEDURE = r'''Private Function BuildRequestJson() As String
    Dim q As String
    q = Chr$(34)
    BuildRequestJson = "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_dir" & q & ":" & q & JsonEscape(ThisWorkbook.Path) & q & "," & _
        q & "overwrite" & q & ":true," & _
        q & "values" & q & ":{" & _
        q & "patient_id" & q & ":" & q & JsonEscape(PatientIdFromSelection()) & q & "," & _
        q & "treatment" & q & ":" & q & JsonEscape(cboTreatment.Value) & q & "," & _
        q & "time" & q & ":" & q & JsonEscape(cboTime.Value) & q & "," & _
        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _
        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & "," & _
        q & "target_base_entry_id" & q & ":" & q & JsonEscape(OutpatientEditTargetBaseEntryId) & q & _
        "}}"
End Function'''

_SUGGESTION_PROCEDURE = r'''Private Function BuildSuggestionRequestJson() As String
    Dim q As String
    Dim modeNumber As Long
    q = Chr$(34)
    modeNumber = cboSuggestMode.ListIndex + 1
    BuildSuggestionRequestJson = "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "values" & q & ":{" & _
        q & "mode" & q & ":" & CStr(modeNumber) & "," & _
        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & "," & _
        q & "time" & q & ":" & q & JsonEscape(cboTime.Value) & q & "," & _
        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _
        q & "target_base_entry_id" & q & ":" & q & JsonEscape(OutpatientEditTargetBaseEntryId) & q & "," & _
        q & "limit" & q & ":5" & _
        "}}"
End Function'''


def _normalize_vba(code: str) -> str:
    return code.replace("\r\n", "\n").replace("\r", "\n")


def _replace_vba_procedure(
    code: str,
    *,
    start_signature: str,
    end_statement: str,
    replacement: str,
) -> str:
    normalized = _normalize_vba(code)
    start = normalized.find(start_signature)
    if start < 0:
        raise RuntimeError(f"Outpatient VBA procedure missing: {start_signature}")
    if normalized.find(start_signature, start + len(start_signature)) >= 0:
        raise RuntimeError(f"Outpatient VBA procedure duplicated: {start_signature}")

    end_marker = "\n" + end_statement
    end = normalized.find(end_marker, start + len(start_signature))
    if end < 0:
        raise RuntimeError(f"Outpatient VBA procedure terminator missing: {start_signature}")
    end += len(end_marker)
    return normalized[:start] + _normalize_vba(replacement) + normalized[end:]


def patch_outpatient_schedule_form_code(code: str) -> str:
    """Add edit-target state without depending on VBIDE line formatting."""

    patched = _normalize_vba(code)

    if 'OutpatientEditTargetBaseEntryId = ""' not in patched:
        signature = "Private Sub UserForm_Initialize()"
        start = patched.find(signature)
        if start < 0:
            raise RuntimeError("Outpatient schedule form initialize procedure missing")
        insert_at = start + len(signature)
        patched = (
            patched[:insert_at]
            + '\n    OutpatientEditTargetBaseEntryId = ""'
            + patched[insert_at:]
        )

    if 'q & "target_base_entry_id" & q' not in patched:
        patched = _replace_vba_procedure(
            patched,
            start_signature="Private Function BuildRequestJson() As String",
            end_statement="End Function",
            replacement=_REQUEST_PROCEDURE,
        )

    if patched.count('q & "target_base_entry_id" & q') < 2:
        patched = _replace_vba_procedure(
            patched,
            start_signature="Private Function BuildSuggestionRequestJson() As String",
            end_statement="End Function",
            replacement=_SUGGESTION_PROCEDURE,
        )

    return patched


def _component_if_present(vbproject, component_name: str):
    try:
        return vbproject.VBComponents(component_name)
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


def install_outpatient_schedule_edit_support(vbproject) -> None:
    """Install edit/preload support while reusing existing VBComponents."""

    form_component = vbproject.VBComponents(FORM_NAME)
    current_form_code = _component_code(form_component)
    patched_form_code = patch_outpatient_schedule_form_code(current_form_code)

    if patched_form_code != _normalize_vba(current_form_code):
        _replace_component_code(form_component, patched_form_code)

    module = _component_if_present(vbproject, EDIT_MODULE_NAME)
    if module is None:
        module = vbproject.VBComponents.Add(1)
        module.Name = EDIT_MODULE_NAME
    _replace_component_code(module, EDIT_MODULE_CODE)

    installed_form_code = _normalize_vba(_component_code(form_component))
    installed_edit_code = _component_code(module)
    if installed_form_code.count('q & "target_base_entry_id" & q') < 2:
        raise RuntimeError("Outpatient edit target was not wired into save and suggestion JSON")
    if "OutpatientEditTargetBaseEntryId" not in installed_form_code:
        raise RuntimeError("Outpatient edit target state was not wired into form")
    if "Public Sub EditOutpatientScheduleByRow(ByVal rowIndex As Long)" not in installed_edit_code:
        raise RuntimeError("Outpatient edit-by-row action was not installed")
    if "ShowOutpatientScheduleManager" not in installed_edit_code:
        raise RuntimeError("Outpatient edit macro was not routed through the manager UI")


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
