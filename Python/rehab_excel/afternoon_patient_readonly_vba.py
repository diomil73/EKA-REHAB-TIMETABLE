"""Optional read-only afternoon panel in the existing patient programme form.

Uses the *existing* PatientProgramRuntime bridge. Runs only during installation
onto a disposable Excel preview. No edit controls, no save, no MASTER changes.
"""
from __future__ import annotations

PANEL_LABEL = "lblAfternoonState"

RUNTIME_MARKER = "Public Function CompletePatientProgram(ByVal patientId As String, ByVal requestArguments As String) As String"
FORM_MARKER = "    LoadProgramConfiguration\nEnd Sub"
FORM_HELPER_MARKER = "Private Sub UserForm_Initialize()"

RUNTIME_INSERT = r'''Public Function LoadPatientAfternoonState(ByVal patientId As String) As String
    LoadPatientAfternoonState = RunProgramBridge( _
        "afternoon-state-vba " & QuoteArg(ThisWorkbook.FullName) & _
        " --patient-id " & QuoteArg(patientId) _
    )
End Function

'''

FORM_HELPER = r'''Private Sub RefreshAfternoonReadOnlyPanel()
    Dim response As String
    Dim lines() As String
    Dim fields() As String
    Dim rowIndex As Long
    Dim detailText As String
    Dim statusText As String
    Dim itemCount As Long

    On Error GoTo PanelError
    lblAfternoonState.Caption = "Απογευματινή ΦΘ: ανάγνωση..."
    response = LoadPatientAfternoonState(mPatientId)
    lines = Split(Replace(response, vbCrLf, vbLf), vbLf)
    statusText = "Απογευματινή ΦΘ: μη διαθέσιμη"
    For rowIndex = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(rowIndex))) = 0 Then GoTo NextLine
        fields = Split(lines(rowIndex), vbTab)
        Select Case fields(0)
            Case "ERROR"
                Err.Raise vbObjectError + 2940, "RefreshAfternoonReadOnlyPanel", _
                    "Η απογευματινή προβολή δεν φορτώθηκε."
            Case "PM_STATUS"
                If UBound(fields) >= 2 Then statusText = fields(2)
            Case "PM_ROW"
                If UBound(fields) >= 3 Then
                    itemCount = itemCount + 1
                    If itemCount <= 2 Then
                        detailText = detailText & vbCrLf & _
                            fields(1) & "  " & fields(2) & "  " & fields(3)
                    End If
                End If
        End Select
NextLine:
    Next rowIndex
    lblAfternoonState.Caption = statusText & detailText
    If itemCount > 2 Then _
        lblAfternoonState.Caption = lblAfternoonState.Caption & _
            vbCrLf & "... και άλλες συνεδρίες."
    Exit Sub
PanelError:
    lblAfternoonState.Caption = "Απογευματινή ΦΘ: η ανάγνωση απέτυχε."
End Sub

'''


def patch_afternoon_panel_text(runtime_code: str, form_code: str) -> tuple[str, str]:
    """Produce source-only patches and refuse unknown/version-drifted VBA."""
    # VBA CodeModule.Lines commonly returns CRLF even though the repository
    # Python string constants use LF. Match source anchors canonically.
    if not isinstance(runtime_code, str) or not isinstance(form_code, str):
        raise RuntimeError("Patient VBA source must be text")
    runtime_code = runtime_code.replace("\r\n", "\n").replace("\r", "\n")
    form_code = form_code.replace("\r\n", "\n").replace("\r", "\n")
    if RUNTIME_INSERT.strip() in runtime_code:
        updated_runtime = runtime_code
    elif runtime_code.count(RUNTIME_MARKER) == 1:
        updated_runtime = runtime_code.replace(RUNTIME_MARKER, RUNTIME_INSERT + RUNTIME_MARKER, 1)
    else:
        raise RuntimeError("Patient runtime marker missing or ambiguous")

    configured = "    LoadProgramConfiguration\n    RefreshAfternoonReadOnlyPanel\nEnd Sub"
    if configured in form_code and FORM_HELPER.strip() in form_code:
        updated_form = form_code
    elif form_code.count(FORM_MARKER) == 1 and form_code.count(FORM_HELPER_MARKER) == 1:
        updated_form = form_code.replace(FORM_MARKER, configured, 1)
        updated_form = updated_form.replace(FORM_HELPER_MARKER, FORM_HELPER + FORM_HELPER_MARKER, 1)
    else:
        raise RuntimeError("Patient form marker missing or ambiguous")
    # The original UserForm_Initialize sets its own height at runtime.
    # Designer.Height alone is insufficient and would clip the PM label.
    if "    Me.Height = 585" in updated_form:
        updated_form = updated_form.replace(
            "    Me.Height = 585", "    Me.Height = 665", 1
        )
    elif "    Me.Height = 665" not in updated_form:
        raise RuntimeError("Patient form runtime height marker missing")
    return updated_runtime, updated_form


def _module_text(component) -> str:
    module = component.CodeModule
    count = int(getattr(module, "CountOfLines", 0) or 0)
    return str(module.Lines(1, count)) if count else ""


def _overwrite_code(component, content: str) -> None:
    module = component.CodeModule
    count = int(getattr(module, "CountOfLines", 0) or 0)
    if count:
        module.DeleteLines(1, count)
    module.AddFromString(content)


def install_afternoon_readonly_panel(vbproject, *, position_control) -> None:
    """Patch only already-installed known components, fail closed on drift."""
    runtime = vbproject.VBComponents("PatientProgramRuntime")
    form = vbproject.VBComponents("frmPatientScheduling")
    runtime_code, form_code = patch_afternoon_panel_text(
        _module_text(runtime), _module_text(form)
    )
    designer = form.Designer
    try:
        label = designer.Controls(PANEL_LABEL)
    except Exception:
        label = designer.Controls.Add("Forms.Label.1", PANEL_LABEL, True)
    label.Caption = "Απογευματινή ΦΘ: προβολή μόνο"
    position_control(label, 620, 525)
    label.Width = 390
    label.Height = 82
    label.WordWrap = True
    try:
        designer.Height = max(float(designer.Height), 640.0)
    except Exception:
        pass
    _overwrite_code(runtime, runtime_code)
    _overwrite_code(form, form_code)
