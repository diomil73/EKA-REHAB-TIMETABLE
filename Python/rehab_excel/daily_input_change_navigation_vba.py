from __future__ import annotations

import unicodedata


MODULE_NAME = "modDailyInputChangeNavigation"
MACRO_NAME = "OpenTherapistDailyWithChanges"
STYLE_MODULE_NAME = "modSupervisorPanelStyle"

MODULE_CODE = r'''Option Explicit

Public Sub OpenTherapistDailyWithChanges()
    ShowTherapistWorkload
    On Error Resume Next
    SyncDailyInputChangesToLedger
    KeepApplicationShell
    StyleSupervisorPanel
    On Error GoTo 0
End Sub
'''

STYLE_MODULE_CODE = r'''Option Explicit

Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const DAILY_SHEET As String = "THERAPIST_DAILY"
Private Const LEDGER_PREFIX As String = "ekaLedger_"
Private Const NAV_MASTER_BUTTON As String = "ekaNav_Master"
Private Const STATUS_ACTIVE As String = "ACTIVE"

Public Sub StyleSupervisorPanel()
    Dim wsDaily As Worksheet
    Dim wsLedger As Worksheet
    Dim startRow As Long
    Dim lastLedgerRow As Long
    Dim ledgerRow As Long
    Dim displayRow As Long
    Dim lastDisplayRow As Long
    Dim actionRow As Long
    Dim shapeIndex As Long
    Dim shapeName As String
    Dim btn As Shape
    Dim nav As Shape
    Dim rowIndex As Long

    On Error GoTo StyleExit
    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)

    startRow = CLng(Val(CStr(wsLedger.Range("R1").Value)))
    If startRow < 2 Then Exit Sub

    lastLedgerRow = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row
    lastDisplayRow = startRow + 2
    For ledgerRow = 2 To lastLedgerRow
        displayRow = CLng(Val(CStr(wsLedger.Cells(ledgerRow, 19).Value)))
        If displayRow > lastDisplayRow Then lastDisplayRow = displayRow
    Next ledgerRow
    actionRow = lastDisplayRow + 2

    ' Repaint the whole management block as one calm application panel.
    With wsDaily.Range("A" & startRow & ":G" & actionRow + 2)
        .Interior.Color = RGB(239, 249, 246)
        .Font.Name = "Calibri"
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With

    With wsDaily.Range("A" & startRow & ":G" & startRow)
        .Interior.Color = RGB(204, 235, 226)
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(38, 86, 76)
        .RowHeight = 30
    End With

    With wsDaily.Range("A" & startRow + 1 & ":G" & startRow + 1)
        .Interior.Color = RGB(250, 226, 229)
        .Font.Bold = True
        .Font.Color = RGB(111, 58, 64)
        .RowHeight = 25
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(222, 194, 198)
        .Borders.Weight = xlThin
    End With

    For rowIndex = startRow + 2 To lastDisplayRow
        wsDaily.Rows(rowIndex).RowHeight = 26
        With wsDaily.Range("A" & rowIndex & ":G" & rowIndex)
            If (rowIndex - startRow) Mod 2 = 0 Then
                .Interior.Color = RGB(248, 253, 251)
            Else
                .Interior.Color = RGB(234, 247, 243)
            End If
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(211, 226, 221)
            .Borders.Weight = xlThin
        End With
    Next rowIndex

    ' Remove the grey legacy form buttons but keep selection checkboxes.
    For shapeIndex = wsDaily.Shapes.Count To 1 Step -1
        shapeName = CStr(wsDaily.Shapes(shapeIndex).Name)
        If Left$(shapeName, Len(LEDGER_PREFIX & "Restore")) = LEDGER_PREFIX & "Restore" Then
            wsDaily.Shapes(shapeIndex).Delete
        End If
    Next shapeIndex

    ' Recreate per-row restore controls as real styled shapes.
    For ledgerRow = 2 To lastLedgerRow
        If UCase$(Trim$(CStr(wsLedger.Cells(ledgerRow, 12).Value))) = STATUS_ACTIVE Then
            displayRow = CLng(Val(CStr(wsLedger.Cells(ledgerRow, 19).Value)))
            If displayRow >= startRow + 2 Then
                Set btn = wsDaily.Shapes.AddShape(5, _
                    wsDaily.Cells(displayRow, 7).Left + 5, _
                    wsDaily.Cells(displayRow, 7).Top + 2, _
                    wsDaily.Cells(displayRow, 7).Width - 10, _
                    wsDaily.Cells(displayRow, 7).Height - 4)
                btn.Name = LEDGER_PREFIX & "Restore_" & CStr(ledgerRow)
                StylePanelButton btn, "Επαναφορά", "RestoreDailyChangeRow", False
            End If
        End If
    Next ledgerRow

    ' One large action bar under the table.
    wsDaily.Rows(actionRow).RowHeight = 38

    Set btn = wsDaily.Shapes.AddShape(5, _
        wsDaily.Cells(actionRow, 3).Left, _
        wsDaily.Cells(actionRow, 3).Top + 2, 190, 32)
    btn.Name = LEDGER_PREFIX & "RestoreSelected"
    StylePanelButton btn, "Επαναφορά επιλεγμένων", "RestoreSelectedDailyChanges", False

    Set btn = wsDaily.Shapes.AddShape(5, _
        wsDaily.Cells(actionRow, 5).Left + 8, _
        wsDaily.Cells(actionRow, 5).Top + 2, 170, 32)
    btn.Name = LEDGER_PREFIX & "RestoreAll"
    StylePanelButton btn, "Επαναφορά όλων", "RestoreAllDailyChanges", True

    On Error Resume Next
    Set nav = wsDaily.Shapes(NAV_MASTER_BUTTON)
    On Error GoTo StyleExit
    If nav Is Nothing Then
        Set nav = wsDaily.Shapes.AddShape(5, _
            wsDaily.Cells(actionRow, 1).Left, _
            wsDaily.Cells(actionRow, 1).Top + 2, 190, 32)
        nav.Name = NAV_MASTER_BUTTON
    Else
        nav.Left = wsDaily.Cells(actionRow, 1).Left
        nav.Top = wsDaily.Cells(actionRow, 1).Top + 2
        nav.Width = 190
        nav.Height = 32
    End If
    StylePanelButton nav, "Επιστροφή στο MASTER", "GoToMasterFromTherapistDaily", False
    nav.Fill.ForeColor.RGB = RGB(202, 235, 225)
    nav.Line.ForeColor.RGB = RGB(104, 164, 148)
    nav.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(36, 82, 70)

    wsDaily.Columns("A").ColumnWidth = 8
    wsDaily.Columns("B").ColumnWidth = 25
    wsDaily.Columns("C:D").ColumnWidth = 23
    wsDaily.Columns("E").ColumnWidth = 30
    wsDaily.Columns("F").ColumnWidth = 18
    wsDaily.Columns("G").ColumnWidth = 22

StyleExit:
End Sub

Private Sub StylePanelButton( _
    ByVal btn As Shape, _
    ByVal caption As String, _
    ByVal macroName As String, _
    ByVal strongAction As Boolean)

    btn.OnAction = macroName
    btn.Placement = 3
    btn.Locked = True
    btn.Line.Weight = 1.25

    If strongAction Then
        btn.Fill.ForeColor.RGB = RGB(241, 196, 199)
        btn.Line.ForeColor.RGB = RGB(192, 113, 121)
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(105, 43, 50)
    Else
        btn.Fill.ForeColor.RGB = RGB(248, 216, 219)
        btn.Line.ForeColor.RGB = RGB(204, 139, 146)
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(101, 50, 56)
    End If

    btn.TextFrame2.TextRange.Text = caption
    btn.TextFrame2.TextRange.Font.Name = "Calibri"
    btn.TextFrame2.TextRange.Font.Size = 10
    btn.TextFrame2.TextRange.Font.Bold = True
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    btn.TextFrame2.VerticalAnchor = 3
End Sub
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _shape_caption(shape) -> str:
    for reader in (
        lambda: shape.TextFrame2.TextRange.Text,
        lambda: shape.TextFrame.Characters().Text,
        lambda: shape.TextFrame.Characters.Text,
    ):
        try:
            value = reader()
        except Exception:
            continue
        if value is not None:
            return str(value).strip()
    return ""


def _fold_control_text(value: object) -> str:
    text = unicodedata.normalize("NFD", str(value or "").strip().casefold())
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _macro_terminal_name(action: object) -> str:
    text = str(action or "").strip()
    terminal = text.rsplit("!", 1)[-1].strip()
    terminal = terminal.rsplit(".", 1)[-1].strip()
    return terminal.casefold()


def _is_legacy_refresh_control(*, name: str, caption: str, macro: str) -> bool:
    name = _fold_control_text(name)
    caption = _fold_control_text(caption)
    macro = _fold_control_text(macro)

    # Keep the original high-confidence identity, but native workbooks have shown
    # variants where Excel rewrites the control name or qualified macro action.
    exact_identity = (
        name == "button 1"
        and caption == "ανανεωση"
        and macro == "refreshall"
    )
    caption_identity = caption == "ανανεωση"
    macro_identity = macro == "refreshall" and name == "button 1"
    return exact_identity or caption_identity or macro_identity


def _remove_legacy_refresh_button(workbook) -> int:
    """Remove the retired THERAPIST_DAILY refresh control across Excel control types."""

    ws = workbook.Worksheets("THERAPIST_DAILY")
    removed = 0

    try:
        buttons = ws.Buttons()
    except Exception:
        buttons = None

    if buttons is not None:
        for index in range(int(buttons.Count), 0, -1):
            button = buttons.Item(index)
            try:
                name = str(button.Name or "")
            except Exception:
                name = ""
            try:
                caption = str(button.Caption or "")
            except Exception:
                caption = ""
            try:
                macro = _macro_terminal_name(button.OnAction)
            except Exception:
                macro = ""
            if _is_legacy_refresh_control(name=name, caption=caption, macro=macro):
                button.Delete()
                removed += 1

    # ActiveX controls are exposed through OLEObjects and may not behave like
    # Forms buttons or ordinary Shapes when reading Caption/OnAction.
    try:
        ole_objects = ws.OLEObjects()
    except Exception:
        ole_objects = None
    if ole_objects is not None:
        for index in range(int(ole_objects.Count), 0, -1):
            ole = ole_objects.Item(index)
            try:
                name = str(ole.Name or "")
            except Exception:
                name = ""
            try:
                caption = str(ole.Object.Caption or "")
            except Exception:
                caption = ""
            if _is_legacy_refresh_control(name=name, caption=caption, macro=""):
                ole.Delete()
                removed += 1

    for index in range(int(ws.Shapes.Count), 0, -1):
        shape = ws.Shapes.Item(index)
        try:
            name = str(shape.Name or "")
        except Exception:
            name = ""
        try:
            macro = _macro_terminal_name(shape.OnAction)
        except Exception:
            macro = ""
        try:
            alt_text = str(shape.AlternativeText or "").strip()
        except Exception:
            alt_text = ""
        caption = _shape_caption(shape).strip() or alt_text
        if _is_legacy_refresh_control(name=name, caption=caption, macro=macro):
            shape.Delete()
            removed += 1

    return removed


def _inject_style_after_refresh(vbproject) -> None:
    """Make every ledger refresh repaint through the application presentation layer."""

    try:
        component = vbproject.VBComponents("modDailyChangeLedger")
    except Exception:
        return
    code_module = component.CodeModule
    source = str(code_module.Lines(1, code_module.CountOfLines))
    marker = '    wsDaily.Columns("G").ColumnWidth = 20\r\n    Exit Sub'
    replacement = '    wsDaily.Columns("G").ColumnWidth = 20\r\n    StyleSupervisorPanel\r\n    Exit Sub'
    if "StyleSupervisorPanel" not in source and marker in source:
        source = source.replace(marker, replacement, 1)
        code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(source)


def _inject_style_after_daily_input_sync(vbproject) -> None:
    """Repaint after DAILY_INPUT rewiring/navigation has finished."""

    try:
        component = vbproject.VBComponents("modDailyInputChangeLedger")
    except Exception:
        return
    code_module = component.CodeModule
    source = str(code_module.Lines(1, code_module.CountOfLines))
    marker = "    RewireDailyInputRestoreButtons ledger\r\n    EnsureTherapistNavigation\r\n    On Error GoTo 0"
    replacement = (
        "    RewireDailyInputRestoreButtons ledger\r\n"
        "    EnsureTherapistNavigation\r\n"
        "    StyleSupervisorPanel\r\n"
        "    On Error GoTo 0"
    )
    if marker in source and "EnsureTherapistNavigation\r\n    StyleSupervisorPanel" not in source:
        source = source.replace(marker, replacement, 1)
        code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(source)


def install_daily_input_change_navigation(vbproject, workbook) -> None:
    """Install supervisor navigation plus the dedicated THERAPIST_DAILY presentation layer."""

    _remove_legacy_refresh_button(workbook)

    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)

    _remove_component_if_present(vbproject, STYLE_MODULE_NAME)
    style_module = vbproject.VBComponents.Add(1)
    style_module.Name = STYLE_MODULE_NAME
    style_module.CodeModule.AddFromString(STYLE_MODULE_CODE)

    _inject_style_after_refresh(vbproject)
    _inject_style_after_daily_input_sync(vbproject)
