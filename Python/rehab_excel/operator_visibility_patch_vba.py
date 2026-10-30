from __future__ import annotations


DAILY_INPUT_MODULE_NAME = "modDailyInputAction"
GLOBAL_REPLACEMENT_MODULE_NAME = "modGlobalReplacementAction"


def _component(vbproject, name: str):
    try:
        return vbproject.VBComponents(name)
    except Exception as exc:
        raise RuntimeError(f"Required VBA component missing: {name}") from exc


def _code(component) -> str:
    module = component.CodeModule
    count = int(module.CountOfLines)
    return str(module.Lines(1, count)) if count else ""


def _replace_all(component, code: str) -> None:
    module = component.CodeModule
    count = int(module.CountOfLines)
    if count:
        module.DeleteLines(1, count)
    module.AddFromString(code)


def patch_daily_input_workload_visibility(code: str) -> str:
    normalized = code.replace("\r\n", "\n").replace("\r", "\n")
    target = '''    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    RenderTherapistWorkloadStrip
    ApplyMasterView'''
    replacement = '''    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    ' Workload is operator-requested information. Do not publish therapist load
    ' badges automatically after every DAILY_INPUT/replacement refresh.
    ApplyMasterView'''
    if target not in normalized:
        raise RuntimeError("DAILY_INPUT workload refresh marker not found")
    return normalized.replace(target, replacement, 1)


_SUMMARY_HELPERS = r'''
Private Function LastReplacementFactRow() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        LastReplacementFactRow = 1
    Else
        LastReplacementFactRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LastReplacementFactRow < 1 Then LastReplacementFactRow = 1
    End If
End Function

Private Function PatientDisplayName(ByVal patientId As String) As String
    Dim ws As Worksheet, lastRow As Long, row As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PATIENTS")
    On Error GoTo 0
    If ws Is Nothing Then
        PatientDisplayName = patientId
        Exit Function
    End If
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = patientId Then
            PatientDisplayName = Trim$(CStr(ws.Cells(row, 3).Value))
            If Len(PatientDisplayName) = 0 Then PatientDisplayName = patientId
            Exit Function
        End If
    Next row
    PatientDisplayName = patientId
End Function

Private Function BuildReplacementSummarySince(ByVal previousLastRow As Long) As String
    Dim ws As Worksheet, lastRow As Long, row As Long
    Dim patientId As String, replacementTherapist As String, replacementTime As String
    Dim statusText As String, result As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = previousLastRow + 1 To lastRow
        statusText = UCase$(Trim$(CStr(ws.Cells(row, 10).Value)))
        If statusText = "ACTIVE" Then
            patientId = Trim$(CStr(ws.Cells(row, 3).Value))
            replacementTherapist = Trim$(CStr(ws.Cells(row, 5).Value))
            replacementTime = Trim$(CStr(ws.Cells(row, 7).Text))
            If Len(replacementTime) = 0 Then replacementTime = Trim$(CStr(ws.Cells(row, 7).Value))
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & PatientDisplayName(patientId) & " -> " & replacementTherapist & " -> " & replacementTime
        End If
    Next row
    BuildReplacementSummarySince = result
End Function
'''


def patch_global_replacement_summary(code: str) -> str:
    normalized = code.replace("\r\n", "\n").replace("\r", "\n")
    normalized = normalized.replace(
        "    Dim exitCode As Long, beforeReview As Long, afterReview As Long",
        "    Dim exitCode As Long, beforeReview As Long, afterReview As Long, previousLastFactRow As Long\n"
        "    Dim replacementSummary As String",
        1,
    )
    marker = '    If Not ConfirmAutomaticReplacementProposal(mode) Then GoTo CleanUp\n    If Not PersistReplacementPayload(payloadPath, mode = "resolved") Then GoTo CleanUp'
    replacement = '    If Not ConfirmAutomaticReplacementProposal(mode) Then GoTo CleanUp\n    previousLastFactRow = LastReplacementFactRow()\n    If Not PersistReplacementPayload(payloadPath, mode = "resolved") Then GoTo CleanUp'
    if marker not in normalized:
        raise RuntimeError("Global replacement persistence marker not found")
    normalized = normalized.replace(marker, replacement, 1)

    success_marker = '''    ResolveDailyReplacementsInPlace = True
    GoTo CleanUp'''
    success_replacement = '''    replacementSummary = BuildReplacementSummarySince(previousLastFactRow)
    If Len(replacementSummary) > 0 Then
        MsgBox "Οι αναπληρώσεις εφαρμόστηκαν:" & vbCrLf & vbCrLf & replacementSummary, _
               vbInformation, "Αναπληρώσεις"
    Else
        MsgBox "Η ημέρα τακτοποιήθηκε χωρίς νέα αναπλήρωση.", vbInformation, "Αναπληρώσεις"
    End If
    ResolveDailyReplacementsInPlace = True
    GoTo CleanUp'''
    if success_marker not in normalized:
        raise RuntimeError("Global replacement success marker not found")
    normalized = normalized.replace(success_marker, success_replacement, 1)

    insert_at = normalized.find("\nPrivate Function ConfirmAutomaticReplacementProposal")
    if insert_at < 0:
        raise RuntimeError("Global replacement helper insertion marker not found")
    normalized = normalized[:insert_at] + "\n" + _SUMMARY_HELPERS + normalized[insert_at:]
    return normalized


def install_operator_visibility_patch(vbproject) -> None:
    daily = _component(vbproject, DAILY_INPUT_MODULE_NAME)
    global_replacement = _component(vbproject, GLOBAL_REPLACEMENT_MODULE_NAME)

    daily_code = patch_daily_input_workload_visibility(_code(daily))
    replacement_code = patch_global_replacement_summary(_code(global_replacement))

    _replace_all(daily, daily_code)
    _replace_all(global_replacement, replacement_code)

    installed_daily = _code(daily)
    installed_replacement = _code(global_replacement)
    if "RenderTherapistWorkloadStrip\n    ApplyMasterView" in installed_daily:
        raise RuntimeError("DAILY_INPUT still auto-renders workload badges")
    if "BuildReplacementSummarySince" not in installed_replacement:
        raise RuntimeError("Replacement summary patch was not installed")

    summary_start = installed_replacement.find("Private Function BuildReplacementSummarySince")
    summary_end = installed_replacement.find("End Function", summary_start)
    if summary_start < 0 or summary_end < 0:
        raise RuntimeError("Replacement summary procedure was not installed")
    summary_code = installed_replacement[summary_start:summary_end]
    required_tokens = (
        "PatientDisplayName(patientId)",
        "replacementTherapist",
        "replacementTime",
    )
    if not all(token in summary_code for token in required_tokens):
        raise RuntimeError("Replacement summary details were not installed")
