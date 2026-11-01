from __future__ import annotations

from .replacement_operator_vba import install_replacement_operator


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


def _normalize_vba(code: str) -> str:
    return code.replace("\r\n", "\n").replace("\r", "\n")


def _replace_all(component, code: str) -> None:
    module = component.CodeModule
    count = int(module.CountOfLines)
    if count:
        module.DeleteLines(1, count)
    module.AddFromString(code)


def patch_daily_input_workload_visibility(code: str) -> str:
    normalized = _normalize_vba(code)
    target = '''    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    RenderTherapistWorkloadStrip
    ApplyMasterView'''
    replacement = '''    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    ' Workload belongs to the therapist card, not the MASTER surface.
    ApplyMasterView'''
    if target not in normalized:
        raise RuntimeError("DAILY_INPUT workload refresh marker not found")
    return normalized.replace(target, replacement, 1)


def patch_daily_input_rich_text_strike_contract(code: str) -> str:
    """Apply and verify MASTER strike state at character level.

    MASTER treatment/status cells may contain mixed formatting. Excel can return
    Null for whole-cell Font.Strikethrough even after a visually correct update,
    so the in-place contract must write and read back every displayed character.
    """

    normalized = _normalize_vba(code)
    if (
        "Private Sub ApplyCharacterStrike(ByVal cell As Range, ByVal enabled As Boolean)" in normalized
        and "Private Function VerifyCharacterStrike(ByVal cell As Range, ByVal enabled As Boolean) As Boolean" in normalized
    ):
        return normalized

    apply_marker = '    If PayloadValue(parts(5)) = "1" Then cell.Font.Strikethrough = True'
    apply_replacement = '    If PayloadValue(parts(5)) = "1" Then ApplyCharacterStrike cell, True'
    if apply_marker not in normalized:
        raise RuntimeError("DAILY_INPUT strike apply marker not found")
    normalized = normalized.replace(apply_marker, apply_replacement, 1)

    verify_marker = '''    If PayloadValue(parts(5)) = "1" And Len(textRuns) = 0 Then
        strikeValue = cell.Font.Strikethrough
        If IsNull(strikeValue) Then Exit Function
        If CBool(strikeValue) <> True Then Exit Function
    End If'''
    verify_replacement = '''    If PayloadValue(parts(5)) = "1" And Len(textRuns) = 0 Then
        If Not VerifyCharacterStrike(cell, True) Then Exit Function
    End If'''
    if verify_marker not in normalized:
        raise RuntimeError("DAILY_INPUT strike verification marker not found")
    normalized = normalized.replace(verify_marker, verify_replacement, 1)

    source_apply_marker = '''        Case "master_daily_non_active_overlay"
            cell.Font.Color = RGB(255, 0, 0)
            cell.Font.Strikethrough = True
        Case "master_daily_patient_status_overlay"
            cell.Font.Color = RGB(255, 0, 0)'''
    source_apply_replacement = '''        Case "master_daily_non_active_overlay"
            cell.Font.Color = RGB(255, 0, 0)
            ApplyCharacterStrike cell, True
        Case "master_daily_patient_status_overlay"
            cell.Font.Color = RGB(255, 0, 0)
            ApplyCharacterStrike cell, False'''
    if source_apply_marker not in normalized:
        raise RuntimeError("MASTER source-tag strike apply marker not found")
    normalized = normalized.replace(source_apply_marker, source_apply_replacement, 1)

    source_verify_marker = '''        Case "master_daily_non_active_overlay"
            strikeValue = cell.Font.Strikethrough
            If IsNull(strikeValue) Then Exit Function
            VerifySourceTagPresentation = _
                (CBool(strikeValue) = True And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
        Case "master_daily_patient_status_overlay"
            VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''
    source_verify_replacement = '''        Case "master_daily_non_active_overlay"
            VerifySourceTagPresentation = _
                (VerifyCharacterStrike(cell, True) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
        Case "master_daily_patient_status_overlay"
            VerifySourceTagPresentation = _
                (VerifyCharacterStrike(cell, False) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''
    if source_verify_marker not in normalized:
        raise RuntimeError("MASTER source-tag strike verification marker not found")
    normalized = normalized.replace(source_verify_marker, source_verify_replacement, 1)

    helper = r'''
Private Sub ApplyCharacterStrike(ByVal cell As Range, ByVal enabled As Boolean)
    Dim textLength As Long
    cell.Font.Strikethrough = enabled
    textLength = Len(CStr(cell.Value))
    If textLength > 0 Then cell.Characters(1, textLength).Font.Strikethrough = enabled
End Sub

Private Function VerifyCharacterStrike(ByVal cell As Range, ByVal enabled As Boolean) As Boolean
    Dim textLength As Long, charIndex As Long
    Dim strikeValue As Variant

    textLength = Len(CStr(cell.Value))
    If textLength = 0 Then
        strikeValue = cell.Font.Strikethrough
        If IsNull(strikeValue) Then Exit Function
        VerifyCharacterStrike = (CBool(strikeValue) = enabled)
        Exit Function
    End If

    For charIndex = 1 To textLength
        strikeValue = cell.Characters(charIndex, 1).Font.Strikethrough
        If IsNull(strikeValue) Then Exit Function
        If CBool(strikeValue) <> enabled Then Exit Function
    Next charIndex
    VerifyCharacterStrike = True
End Function
'''
    insert_at = normalized.find("\nPrivate Sub ApplySourceTagPresentation")
    if insert_at < 0:
        raise RuntimeError("MASTER character-strike helper insertion marker not found")
    return normalized[:insert_at] + "\n" + helper + normalized[insert_at:]


def patch_daily_input_patient_status_visibility(code: str) -> str:
    """Keep MASTER status/reason readable after both patching and view refresh."""

    normalized = _normalize_vba(code)
    canonical_rich_text = (
        "ApplyCharacterStrike cell, False" in normalized
        and "VerifyCharacterStrike(cell, False)" in normalized
    )

    if not canonical_rich_text:
        apply_marker = '''        Case "master_daily_patient_status_overlay"
            cell.Font.Color = RGB(255, 0, 0)'''
        apply_replacement = '''        Case "master_daily_patient_status_overlay"
            cell.Font.Color = RGB(255, 0, 0)
            cell.Font.Strikethrough = False'''
        if apply_marker not in normalized:
            raise RuntimeError("MASTER patient-status presentation marker not found")
        normalized = normalized.replace(apply_marker, apply_replacement, 1)

        verify_marker = '''        Case "master_daily_patient_status_overlay"
            VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''
        verify_replacement = '''        Case "master_daily_patient_status_overlay"
            strikeValue = cell.Font.Strikethrough
            If IsNull(strikeValue) Then Exit Function
            VerifySourceTagPresentation = _
                (CBool(strikeValue) = False And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''
        if verify_marker not in normalized:
            raise RuntimeError("MASTER patient-status verification marker not found")
        normalized = normalized.replace(verify_marker, verify_replacement, 1)

    refresh_marker = '''    ApplyMasterView
    RefreshOperationalViews = True'''
    refresh_replacement = '''    ApplyMasterView
    NormalizeMasterPatientStatusPresentation master
    RefreshOperationalViews = True'''
    if refresh_marker not in normalized:
        raise RuntimeError("MASTER view refresh marker not found")
    normalized = normalized.replace(refresh_marker, refresh_replacement, 1)

    helper = r'''
Private Sub NormalizeMasterPatientStatusPresentation(ByVal master As Worksheet)
    Dim statusCol As Long, headerRow As Long, lastRow As Long, col As Long, row As Long
    Dim roomText As String, patientText As String, headerText As String

    For row = 1 To Application.WorksheetFunction.Min(master.UsedRange.Rows.Count, 12)
        roomText = LCase$(Trim$(CStr(master.Cells(row, 2).Value)))
        patientText = LCase$(Trim$(CStr(master.Cells(row, 3).Value)))
        If roomText = LCase$("Θάλαμος") And patientText = LCase$("Ασθενής") Then
            headerRow = row
            Exit For
        End If
    Next row
    If headerRow = 0 Then Exit Sub

    For col = 1 To Application.WorksheetFunction.Min(master.Cells(headerRow, master.Columns.Count).End(xlToLeft).Column, 20)
        headerText = LCase$(Trim$(CStr(master.Cells(headerRow, col).Value)))
        If headerText = LCase$("Κατάσταση") Then
            statusCol = col
            Exit For
        End If
    Next col
    If statusCol = 0 Then Exit Sub

    lastRow = master.Cells(master.Rows.Count, statusCol).End(xlUp).Row
    If lastRow <= headerRow Then Exit Sub
    For row = headerRow + 1 To lastRow
        If Len(Trim$(CStr(master.Cells(row, statusCol).Value))) > 0 Then
            ApplyCharacterStrike master.Cells(row, statusCol), False
        End If
    Next row
End Sub
'''
    insert_at = normalized.find("\nPrivate Sub ValidatePatchTarget")
    if insert_at < 0:
        raise RuntimeError("MASTER status helper insertion marker not found")
    return normalized[:insert_at] + "\n" + helper + normalized[insert_at:]


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
    If ws Is Nothing Then PatientDisplayName = patientId: Exit Function
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
    normalized = _normalize_vba(code)
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

    old_prompt = '''             "Να εφαρμοστεί τώρα η αυτόματη πρόταση;" & vbCrLf & _
             "Επιλέξτε Όχι αν θέλετε να κάνετε χειροκίνητες επιλογές ασθενή-ασθενή."'''
    new_prompt = '''             "Να εφαρμοστεί τώρα η αυτόματη πρόταση;" & vbCrLf & _
             "Επιλέξτε Όχι για επιστροφή χωρίς εφαρμογή της πρότασης."'''
    if old_prompt not in normalized:
        raise RuntimeError("Global replacement confirmation prompt marker not found")
    normalized = normalized.replace(old_prompt, new_prompt, 1)

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
    return normalized[:insert_at] + "\n" + _SUMMARY_HELPERS + normalized[insert_at:]


def install_operator_visibility_patch(vbproject) -> None:
    daily = _component(vbproject, DAILY_INPUT_MODULE_NAME)
    global_replacement = _component(vbproject, GLOBAL_REPLACEMENT_MODULE_NAME)

    daily_code = patch_daily_input_workload_visibility(_code(daily))
    daily_code = patch_daily_input_rich_text_strike_contract(daily_code)
    daily_code = patch_daily_input_patient_status_visibility(daily_code)
    _replace_all(daily, daily_code)
    _replace_all(global_replacement, patch_global_replacement_summary(_code(global_replacement)))

    install_replacement_operator(vbproject)

    installed_daily = _normalize_vba(_code(daily))
    installed_replacement = _normalize_vba(_code(global_replacement))
    if "RenderTherapistWorkloadStrip\n    ApplyMasterView" in installed_daily:
        raise RuntimeError("DAILY_INPUT still auto-renders workload badges")
    if "ResolveReplacementWorkflow()" not in installed_daily:
        raise RuntimeError("DAILY_INPUT was not routed through the replacement operator")
    if "ApplyCharacterStrike cell, False" not in installed_daily:
        raise RuntimeError("MASTER patient status still permits character-level strike-through")
    if "VerifyCharacterStrike(cell, True)" not in installed_daily:
        raise RuntimeError("MASTER treatment strike read-back is not character-safe")
    if "NormalizeMasterPatientStatusPresentation master" not in installed_daily:
        raise RuntimeError("MASTER status normalization was not installed after ApplyMasterView")
    if "ApplyCharacterStrike master.Cells(row, statusCol), False" not in installed_daily:
        raise RuntimeError("MASTER status normalization is not character-safe")
    if "BuildReplacementSummarySince" not in installed_replacement:
        raise RuntimeError("Replacement summary patch was not installed")
    if "χειροκίνητες επιλογές ασθενή-ασθενή" in installed_replacement:
        raise RuntimeError("Global automatic prompt still claims to own manual routing")

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
