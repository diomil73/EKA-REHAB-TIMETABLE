from __future__ import annotations


IN_SESSION_MODULE_NAME = "modOutpatientInSession"

_EDIT_REMOVE_CALL = (
    "        RemoveFromTherapistDailyForCurrentDate patientName, oldTherapistName, oldTimeText, oldDayPattern"
)

_EDIT_REMOVE_REPLACEMENT = r'''        RemoveOutpatientRenderFromAffectedDailyCells _
            patientName, oldTherapistName, oldTimeText, oldDayPattern, _
            therapistName, timeText, dayPattern'''

_REFRESH_CALL = (
    "    RefreshTherapistDailyForCurrentDate patientName, therapistName, timeText, dayPattern, treatment"
)

_REFRESH_REPLACEMENT = r'''    RefreshTherapistDailyForCurrentDate patientName, therapistName, timeText, dayPattern, treatment
    On Error Resume Next
    RenderTherapistWorkloadStrip
    On Error GoTo SaveError'''

_HELPERS = r'''
Private Sub RemoveOutpatientRenderFromAffectedDailyCells( _
    ByVal patientName As String, _
    ByVal oldTherapistName As String, _
    ByVal oldTimeText As String, _
    ByVal oldDayPattern As String, _
    ByVal newTherapistName As String, _
    ByVal newTimeText As String, _
    ByVal newDayPattern As String _
)
    Dim dailyInput As Worksheet
    Dim dailySheet As Worksheet
    Dim targetDate As Date

    On Error GoTo RemoveError

    Set dailyInput = ThisWorkbook.Worksheets("DAILY_INPUT")
    If Not TryReadDailyInputDate(dailyInput.Range("B2").Value, targetDate) Then Exit Sub
    Set dailySheet = ThisWorkbook.Worksheets("THERAPIST_DAILY")

    If DayPatternIncludesDate(oldDayPattern, targetDate) Then
        RemovePatientRenderFromDailyCell dailySheet, patientName, oldTherapistName, oldTimeText
    End If

    If DayPatternIncludesDate(newDayPattern, targetDate) Then
        If StrComp(NormalizeProviderName(oldTherapistName), NormalizeProviderName(newTherapistName), vbTextCompare) <> 0 _
           Or StrComp(TimeKey(oldTimeText), TimeKey(newTimeText), vbTextCompare) <> 0 Then
            RemovePatientRenderFromDailyCell dailySheet, patientName, newTherapistName, newTimeText
        End If
    End If
    Exit Sub

RemoveError:
    Err.Raise Err.Number, "RemoveOutpatientRenderFromAffectedDailyCells", Err.Description
End Sub

Private Sub RemovePatientRenderFromDailyCell( _
    ByVal dailySheet As Worksheet, _
    ByVal patientName As String, _
    ByVal therapistName As String, _
    ByVal timeText As String _
)
    Dim targetCell As Range
    Dim normalizedTime As Date
    Dim originalText As String
    Dim lines() As String
    Dim kept As String
    Dim index As Long
    Dim lineText As String

    If Len(Trim$(therapistName)) = 0 Or Len(Trim$(timeText)) = 0 Then Exit Sub
    If Not IsDate(Replace(timeText, ".", ":")) Then Exit Sub

    normalizedTime = TimeValue(Replace(timeText, ".", ":"))
    Set targetCell = FindTherapistDailyCell(dailySheet, therapistName, normalizedTime)
    If targetCell Is Nothing Then Exit Sub

    originalText = CStr(targetCell.Value)
    If Len(originalText) = 0 Then Exit Sub

    lines = Split(Replace(originalText, vbCrLf, vbLf), vbLf)
    kept = ""
    For index = LBound(lines) To UBound(lines)
        lineText = CStr(lines(index))
        If Not IsOutpatientPatientRenderLine(lineText, patientName) Then
            If Len(kept) > 0 Then kept = kept & vbLf
            kept = kept & lineText
        End If
    Next index

    targetCell.Value = kept
    targetCell.WrapText = True
    If Len(Trim$(kept)) = 0 Then targetCell.Interior.Pattern = xlNone
End Sub

Private Function IsOutpatientPatientRenderLine( _
    ByVal lineText As String, _
    ByVal patientName As String _
) As Boolean
    Dim text As String
    Dim suffix As String

    text = Trim$(lineText)
    patientName = Trim$(patientName)
    If Len(text) = 0 Or Len(patientName) = 0 Then Exit Function

    If StrComp(text, patientName, vbTextCompare) = 0 Then
        IsOutpatientPatientRenderLine = True
        Exit Function
    End If

    If Len(text) <= Len(patientName) Then Exit Function
    If StrComp(Left$(text, Len(patientName)), patientName, vbTextCompare) <> 0 Then Exit Function

    suffix = Mid$(text, Len(patientName) + 1)
    If Left$(suffix, 2) = " [" _
       Or Left$(suffix, 2) = " (" _
       Or Left$(suffix, 2) = " {" _
       Or Left$(suffix, 3) = " → " Then
        IsOutpatientPatientRenderLine = True
    End If
End Function
'''


def patch_outpatient_in_session_module_code(code: str) -> str:
    """Harden outpatient edit cleanup and refresh the visible workload strip.

    The legacy in-session writer removed only a line exactly equal to the patient
    display name. The daily renderer may include a suffix such as ``[Καθ/να]``;
    an edit then left the old rendered line behind and appended a second plain
    patient line. This patch cleans both the old slot and the destination slot
    using the accepted rendered-line shapes before the writer renders once.

    After the daily projection is updated, the workload strip is recalculated so
    the MASTER does not keep stale therapist-load badges after an edit. The call
    is optional at runtime so workbooks without the strip module still save.
    """

    normalized = code.replace("\r\n", "\n").replace("\r", "\n")

    if "Private Sub RemoveOutpatientRenderFromAffectedDailyCells(" not in normalized:
        if normalized.count(_EDIT_REMOVE_CALL) != 1:
            raise RuntimeError("Expected exactly one outpatient edit daily-removal call")
        normalized = normalized.replace(_EDIT_REMOVE_CALL, _EDIT_REMOVE_REPLACEMENT, 1)
        marker = "\nPrivate Function TryReadDailyInputDate"
        if marker not in normalized:
            raise RuntimeError("Could not locate outpatient in-session helper insertion point")
        normalized = normalized.replace(marker, _HELPERS + marker, 1)

    if "RenderTherapistWorkloadStrip" not in normalized:
        if normalized.count(_REFRESH_CALL) != 1:
            raise RuntimeError("Expected exactly one outpatient daily refresh call")
        normalized = normalized.replace(_REFRESH_CALL, _REFRESH_REPLACEMENT, 1)

    return normalized


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


def install_outpatient_daily_edit_hardening(vbproject) -> None:
    component = vbproject.VBComponents(IN_SESSION_MODULE_NAME)
    current = _component_code(component)
    patched = patch_outpatient_in_session_module_code(current)
    if patched != current.replace("\r\n", "\n").replace("\r", "\n"):
        _replace_component_code(component, patched)

    installed = _component_code(component)
    if "RemoveOutpatientRenderFromAffectedDailyCells" not in installed:
        raise RuntimeError("Outpatient daily edit hardening was not installed")
    if "IsOutpatientPatientRenderLine" not in installed:
        raise RuntimeError("Outpatient rendered-line cleanup was not installed")
    if "RenderTherapistWorkloadStrip" not in installed:
        raise RuntimeError("Outpatient workload refresh hook was not installed")
