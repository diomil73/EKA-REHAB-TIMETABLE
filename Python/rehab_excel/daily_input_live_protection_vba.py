from __future__ import annotations


MODULE_NAME = "modDailyInputLiveProtection"
DAILY_INPUT_LEDGER_MODULE = "modDailyInputChangeLedger"
REPLACEMENT_OPERATOR_MODULE = "modReplacementOperator"

MODULE_CODE = r'''Option Explicit

Private Const DAILY_INPUT_SHEET As String = "DAILY_INPUT"
Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const META_PREFIX As String = "DAILY_INPUT:"
Private Const STATUS_ACTIVE As String = "ACTIVE"

Public Function RefreshDailyInputProtectionFromLedger() As Boolean
    Dim ws As Worksheet, ledger As Worksheet
    Dim lastRow As Long, row As Long, sourceRow As Long
    Dim meta As String, sourceRowText As String

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(DAILY_INPUT_SHEET)

    On Error Resume Next
    Set ledger = ThisWorkbook.Worksheets(LEDGER_SHEET)
    On Error GoTo Failed

    ws.Unprotect
    ws.Range("A7:F16").Locked = False
    ws.Range("A21:E40").Locked = False

    If Not ledger Is Nothing Then
        lastRow = ledger.Cells(ledger.Rows.Count, 1).End(xlUp).Row
        For row = 2 To lastRow
            If UCase$(Trim$(CStr(ledger.Cells(row, 12).Value))) = STATUS_ACTIVE Then
                meta = Trim$(CStr(ledger.Cells(row, 18).Value))
                If Left$(meta, Len(META_PREFIX)) = META_PREFIX Then
                    sourceRowText = Mid$(meta, InStrRev(meta, ":") + 1)
                    If Not IsNumeric(sourceRowText) Then Err.Raise vbObjectError + 2561, , "Invalid ACTIVE DAILY_INPUT source row"
                    sourceRow = CLng(sourceRowText)
                    If sourceRow >= 7 And sourceRow <= 16 Then
                        ws.Range("A" & CStr(sourceRow) & ":F" & CStr(sourceRow)).Locked = True
                    ElseIf sourceRow >= 21 And sourceRow <= 40 Then
                        ws.Range("A" & CStr(sourceRow) & ":E" & CStr(sourceRow)).Locked = True
                    Else
                        Err.Raise vbObjectError + 2562, , "ACTIVE DAILY_INPUT source row outside supported input ranges"
                    End If
                End If
            End If
        Next row
    End If

    ' UserInterfaceOnly is session-scoped in Excel, so re-apply it whenever the live workflow runs.
    ws.Protect UserInterfaceOnly:=True
    RefreshDailyInputProtectionFromLedger = True
    Exit Function

Failed:
    WriteProtectionDiagnostic Err.Number, Err.Description, "refresh protection"
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Function

Private Sub WriteProtectionDiagnostic(ByVal errorNumber As Long, ByVal errorText As String, ByVal phase As String)
    Dim filePath As String, handle As Integer
    filePath = Environ$("TEMP") & "\\eka_daily_input_protection_last_error.log"
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #handle, "phase=" & phase
    Print #handle, "error_number=" & CStr(errorNumber)
    Print #handle, "error=" & errorText
    Close #handle
End Sub
'''


_LEDGER_MARKER = (
    '    ledger.Visible = xlSheetVeryHidden\r\n'
    '    ThisWorkbook.Save'
)
_LEDGER_REPLACEMENT = (
    '    ledger.Visible = xlSheetVeryHidden\r\n'
    '    If Not RefreshDailyInputProtectionFromLedger() Then _\r\n'
    '        Err.Raise vbObjectError + 2563, , "DAILY_INPUT protection refresh failed"\r\n'
    '    ThisWorkbook.Save'
)

_OPERATOR_MARKER = (
    '    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo RefreshExit'
)
_OPERATOR_REPLACEMENT = (
    '    If Not RefreshDailyInputProtectionFromLedger() Then\r\n'
    '        WriteOperatorDiagnostic vbObjectError + 2564, "DAILY_INPUT protection refresh failed", "operator projection protection"\r\n'
    '        GoTo RefreshExit\r\n'
    '    End If\r\n'
    '    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo RefreshExit'
)


def _component_if_present(vbproject, name: str):
    try:
        return vbproject.VBComponents(name)
    except Exception:
        return None


def _replace_exact_once(code: str, marker: str, replacement: str, *, label: str) -> str:
    if replacement in code:
        return code
    count = code.count(marker)
    if count != 1:
        raise RuntimeError(f"{label}: expected exact marker once, found {count}")
    return code.replace(marker, replacement, 1)


def _patch_component(vbproject, name: str, marker: str, replacement: str, *, label: str) -> None:
    component = _component_if_present(vbproject, name)
    if component is None:
        raise RuntimeError(f"{label}: missing VBA component {name}")
    module = component.CodeModule
    code = str(module.Lines(1, module.CountOfLines))
    patched = _replace_exact_once(code, marker, replacement, label=label)
    if patched != code:
        module.DeleteLines(1, module.CountOfLines)
        module.AddFromString(patched)


def install_daily_input_live_protection(vbproject) -> None:
    existing = _component_if_present(vbproject, MODULE_NAME)
    if existing is not None:
        vbproject.VBComponents.Remove(existing)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)

    _patch_component(
        vbproject,
        DAILY_INPUT_LEDGER_MODULE,
        _LEDGER_MARKER,
        _LEDGER_REPLACEMENT,
        label="DAILY_INPUT ledger live-protection hook",
    )
    _patch_component(
        vbproject,
        REPLACEMENT_OPERATOR_MODULE,
        _OPERATOR_MARKER,
        _OPERATOR_REPLACEMENT,
        label="replacement operator live-protection hook",
    )
