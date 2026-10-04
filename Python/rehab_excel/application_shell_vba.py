from __future__ import annotations

APP_SHELL_MODULE_NAME = "modApplicationShell"

APP_SHELL_MODULE_CODE = r'''Option Explicit

Private shellStateCaptured As Boolean
Private previousFormulaBar As Boolean
Private previousStatusBar As Boolean
Private previousRibbonVisible As Boolean

Public Sub EnterApplicationShell()
    On Error GoTo SafeExit

    If Not shellStateCaptured Then
        previousFormulaBar = Application.DisplayFormulaBar
        previousStatusBar = Application.DisplayStatusBar
        previousRibbonVisible = RibbonIsVisible()
        shellStateCaptured = True
    End If

    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    SetRibbonVisible False

    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
    End If

    If WorksheetExists("MASTER_SCHEDULE") Then
        ThisWorkbook.Worksheets("MASTER_SCHEDULE").Activate
        ThisWorkbook.Worksheets("MASTER_SCHEDULE").Range("A1").Select
    End If

SafeExit:
End Sub

Public Sub ExitApplicationShell()
    On Error Resume Next

    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayWorkbookTabs = True
        ActiveWindow.DisplayHeadings = True
        ActiveWindow.DisplayGridlines = True
    End If

    If shellStateCaptured Then
        Application.DisplayFormulaBar = previousFormulaBar
        Application.DisplayStatusBar = previousStatusBar
        SetRibbonVisible previousRibbonVisible
    Else
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        SetRibbonVisible True
    End If

    On Error GoTo 0
End Sub

Public Sub RestoreExcelInterface()
    ExitApplicationShell
    shellStateCaptured = False
End Sub

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function RibbonIsVisible() As Boolean
    On Error GoTo AssumeVisible
    RibbonIsVisible = Application.CommandBars("Ribbon").Visible
    Exit Function
AssumeVisible:
    RibbonIsVisible = True
End Function

Private Sub SetRibbonVisible(ByVal makeVisible As Boolean)
    On Error Resume Next
    If makeVisible Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    End If
    On Error GoTo 0
End Sub
'''


_EVENT_SPECS = (
    ("Workbook_Open", "Private Sub Workbook_Open()", "EnterApplicationShell"),
    ("Workbook_Activate", "Private Sub Workbook_Activate()", "EnterApplicationShell"),
    ("Workbook_Deactivate", "Private Sub Workbook_Deactivate()", "ExitApplicationShell"),
    (
        "Workbook_BeforeClose",
        "Private Sub Workbook_BeforeClose(Cancel As Boolean)",
        "ExitApplicationShell",
    ),
)


def _procedure_bounds(code_module, procedure_name: str) -> tuple[int, int] | None:
    try:
        start = int(code_module.ProcStartLine(procedure_name, 0))
        count = int(code_module.ProcCountLines(procedure_name, 0))
    except Exception:
        return None
    if start < 1 or count < 1:
        return None
    return start, count


def _procedure_text(code_module, procedure_name: str) -> str:
    bounds = _procedure_bounds(code_module, procedure_name)
    if bounds is None:
        return ""
    start, count = bounds
    return str(code_module.Lines(start, count))


def _ensure_event_call(
    code_module,
    *,
    procedure_name: str,
    signature: str,
    call_name: str,
) -> None:
    existing = _procedure_text(code_module, procedure_name)
    if existing:
        if call_name.casefold() in existing.casefold():
            return
        start, _ = _procedure_bounds(code_module, procedure_name) or (0, 0)
        code_module.InsertLines(start + 1, f"    {call_name}")
        return

    if int(code_module.CountOfLines) > 0:
        code_module.InsertLines(int(code_module.CountOfLines) + 1, "")
    code_module.InsertLines(
        int(code_module.CountOfLines) + 1,
        signature + "\r\n    " + call_name + "\r\nEnd Sub",
    )


def _workbook_document_component(vbproject, workbook=None):
    if workbook is not None:
        code_name = str(getattr(workbook, "CodeName", "") or "").strip()
        if code_name:
            try:
                return vbproject.VBComponents(code_name)
            except Exception:
                pass

    # Fallback for localized/unusual projects. Worksheet and workbook modules
    # are all document components (type 100), but only the workbook document
    # exposes workbook-level event procedures such as Workbook_Open.
    for index in range(1, int(vbproject.VBComponents.Count) + 1):
        component = vbproject.VBComponents(index)
        try:
            if int(component.Type) != 100:
                continue
        except Exception:
            continue
        name = str(getattr(component, "Name", "") or "")
        if name.casefold() == "thisworkbook":
            return component

    raise RuntimeError("Workbook VBA document module could not be resolved")


def install_application_shell(vbproject, *, workbook=None) -> None:
    try:
        existing = vbproject.VBComponents(APP_SHELL_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = APP_SHELL_MODULE_NAME
    module.CodeModule.AddFromString(APP_SHELL_MODULE_CODE)

    this_workbook = _workbook_document_component(vbproject, workbook)
    code_module = this_workbook.CodeModule
    for procedure_name, signature, call_name in _EVENT_SPECS:
        _ensure_event_call(
            code_module,
            procedure_name=procedure_name,
            signature=signature,
            call_name=call_name,
        )
