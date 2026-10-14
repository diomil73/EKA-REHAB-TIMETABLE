from __future__ import annotations

APP_SHELL_MODULE_NAME = "modApplicationShell"
MASTER_SIDEBAR_FORM_NAME = "frmMasterSidebar"

MASTER_SIDEBAR_FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "EKA"
        .Width = 180
        .Height = 250
        .StartUpPosition = 0
        .BackColor = RGB(245, 247, 250)
    End With

    StyleSidebarButton cmdRegistry, "Διαχείριση Μητρώων", 24
    StyleSidebarButton cmdAbsences, "Απουσίες / Ακυρώσεις", 70
    StyleSidebarButton cmdDaily, "Ημερήσιο πρόγραμμα", 116
    StyleSidebarButton cmdExit, "Save & Exit", 174
End Sub

Private Sub StyleSidebarButton(ByVal button As MSForms.CommandButton, ByVal text As String, ByVal topPosition As Single)
    With button
        .Caption = text
        .Left = 18
        .Top = topPosition
        .Width = 138
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub cmdRegistry_Click()
    ToolbarRegistryManagement
End Sub

Private Sub cmdAbsences_Click()
    ToolbarAbsences
End Sub

Private Sub cmdDaily_Click()
    ToolbarTherapistDaily
End Sub

Private Sub cmdExit_Click()
    ExitApplication
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub
'''

APP_SHELL_MODULE_CODE = r'''Option Explicit

Private shellStateCaptured As Boolean
Private previousFormulaBar As Boolean
Private previousStatusBar As Boolean
Private previousRibbonVisible As Boolean

Public Sub EnterApplicationShell()
    On Error GoTo SafeExit

    ApplyApplicationChrome

    If Not ActivatePostCommitTarget() Then
        If WorksheetExists("MASTER_SCHEDULE") Then
            ThisWorkbook.Worksheets("MASTER_SCHEDULE").Activate
            ApplyMasterView
            ShowMasterSidebar
            ThisWorkbook.Worksheets("MASTER_SCHEDULE").Range("A1").Select
        End If
    End If

SafeExit:
End Sub

Public Sub KeepApplicationShell()
    On Error Resume Next
    ApplyApplicationChrome
    If ActiveSheet.Name = "MASTER_SCHEDULE" Then
        ApplyMasterView
        ShowMasterSidebar
    Else
        HideMasterSidebar
    End If
    On Error GoTo 0
End Sub

Public Sub ShowMasterSidebar()
    On Error Resume Next
    If ActiveSheet.Name <> "MASTER_SCHEDULE" Then Exit Sub

    With frmMasterSidebar
        .StartUpPosition = 0
        .Left = Application.Left + 4
        .Top = Application.Top + 72
        If Not .Visible Then .Show vbModeless
    End With

    SetMasterToolbarShapesVisible False
    On Error GoTo 0
End Sub

Public Sub HideMasterSidebar()
    On Error Resume Next
    frmMasterSidebar.Hide
    SetMasterToolbarShapesVisible True
    On Error GoTo 0
End Sub

Private Sub SetMasterToolbarShapesVisible(ByVal makeVisible As Boolean)
    Dim ws As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    If ws Is Nothing Then Exit Sub

    For Each shp In ws.Shapes
        If Left$(shp.Name, 11) = "ekaToolbar_" Then
            shp.Visible = makeVisible
        End If
    Next shp
    On Error GoTo 0
End Sub

Private Sub ApplyApplicationChrome()
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
End Sub

Public Sub ExitApplicationShell()
    On Error Resume Next

    HideMasterSidebar
    Unload frmMasterSidebar

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

Private Function ActivatePostCommitTarget() As Boolean
    Dim markerPath As String
    Dim targetName As String
    Dim fileNumber As Integer

    On Error GoTo NoTarget
    markerPath = ThisWorkbook.FullName & ".eka_next_sheet"
    If Dir$(markerPath) = "" Then Exit Function

    fileNumber = FreeFile
    Open markerPath For Input As #fileNumber
    Line Input #fileNumber, targetName
    Close #fileNumber
    Kill markerPath

    targetName = Trim$(targetName)
    If Len(targetName) = 0 Then Exit Function
    If Not WorksheetExists(targetName) Then Exit Function

    ThisWorkbook.Worksheets(targetName).Activate
    KeepApplicationShell
    ActivatePostCommitTarget = True
    Exit Function

NoTarget:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    If Len(markerPath) > 0 And Dir$(markerPath) <> "" Then Kill markerPath
    On Error GoTo 0
End Function

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
        "Workbook_SheetActivate",
        "Private Sub Workbook_SheetActivate(ByVal Sh As Object)",
        "KeepApplicationShell",
    ),
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


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _install_master_sidebar_form(vbproject) -> None:
    _remove_component_if_present(vbproject, MASTER_SIDEBAR_FORM_NAME)
    form = vbproject.VBComponents.Add(3)
    form.Name = MASTER_SIDEBAR_FORM_NAME
    designer = form.Designer
    try:
        designer.Caption = "EKA"
    except Exception:
        pass

    for name, caption, top in (
        ("cmdRegistry", "Διαχείριση Μητρώων", 24),
        ("cmdAbsences", "Απουσίες / Ακυρώσεις", 70),
        ("cmdDaily", "Ημερήσιο πρόγραμμα", 116),
        ("cmdExit", "Save & Exit", 174),
    ):
        button = designer.Controls.Add("Forms.CommandButton.1", name, True)
        button.Caption = caption
        button.Left = 18
        button.Top = top
        button.Width = 138
        button.Height = 34

    form.CodeModule.AddFromString(MASTER_SIDEBAR_FORM_CODE)


def install_application_shell(vbproject, *, workbook=None) -> None:
    _remove_component_if_present(vbproject, APP_SHELL_MODULE_NAME)

    module = vbproject.VBComponents.Add(1)
    module.Name = APP_SHELL_MODULE_NAME
    module.CodeModule.AddFromString(APP_SHELL_MODULE_CODE)

    _install_master_sidebar_form(vbproject)

    this_workbook = _workbook_document_component(vbproject, workbook)
    code_module = this_workbook.CodeModule
    for procedure_name, signature, call_name in _EVENT_SPECS:
        _ensure_event_call(
            code_module,
            procedure_name=procedure_name,
            signature=signature,
            call_name=call_name,
        )
