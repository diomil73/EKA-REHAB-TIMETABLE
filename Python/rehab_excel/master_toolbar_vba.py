from __future__ import annotations

MASTER_TOOLBAR_MODULE_NAME = "modMasterToolbar"
MASTER_TOOLBAR_PREFIX = "ekaToolbar_"

MASTER_TOOLBAR_MODULE_CODE = r'''Option Explicit

Public Sub ToolbarNewPatient()
    frmNewPatient.Show
End Sub

Public Sub ToolbarNewProvider()
    ShowRegistrationMenu
End Sub

Public Sub ToolbarPatientAbsence()
    GoToDailyInputSection "ΑΠΟΥΣΙΕΣ / ΑΚΥΡΩΣΕΙΣ ΑΣΘΕΝΩΝ"
End Sub

Public Sub ToolbarTherapistAbsence()
    GoToDailyInputSection "ΑΠΟΥΣΙΕΣ ΘΕΡΑΠΕΥΤΩΝ"
End Sub

Public Sub GoToMaster()
    On Error GoTo MissingSheet
    ThisWorkbook.Worksheets("MASTER_SCHEDULE").Activate
    ThisWorkbook.Worksheets("MASTER_SCHEDULE").Range("A1").Select
    KeepApplicationShell
    Exit Sub
MissingSheet:
    MsgBox "Δεν βρέθηκε το MASTER_SCHEDULE.", vbExclamation, "Κεντρική οθόνη"
End Sub

Public Sub ExitApplication()
    Dim eventsWereEnabled As Boolean

    On Error GoTo ExitError

    ThisWorkbook.Save

    On Error Resume Next
    RestoreExcelInterface
    On Error GoTo ExitError

    eventsWereEnabled = Application.EnableEvents
    Application.EnableEvents = False
    ThisWorkbook.Close SaveChanges:=False
    Application.EnableEvents = eventsWereEnabled
    Exit Sub

ExitError:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "Δεν ήταν δυνατή η αποθήκευση και έξοδος: " & Err.Description, _
           vbExclamation, "Save & Exit"
End Sub

Public Sub ToolbarTherapistDaily()
    On Error GoTo MissingSheet
    ThisWorkbook.Worksheets("THERAPIST_DAILY").Activate
    Exit Sub
MissingSheet:
    MsgBox "Δεν βρέθηκε το THERAPIST_DAILY.", vbExclamation, "Ημερήσιο πρόγραμμα"
End Sub

Private Sub GoToDailyInputSection(ByVal sectionTitle As String)
    Dim ws As Worksheet
    Dim found As Range

    On Error GoTo MissingSheet
    Set ws = ThisWorkbook.Worksheets("DAILY_INPUT")
    ws.Activate

    Set found = ws.Cells.Find( _
        What:=sectionTitle, _
        After:=ws.Cells(1, 1), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If Not found Is Nothing Then
        Application.Goto found, True
    Else
        ws.Range("A1").Select
    End If
    Exit Sub

MissingSheet:
    MsgBox "Δεν βρέθηκε το DAILY_INPUT.", vbExclamation, "Ημερήσια κατάσταση"
End Sub
'''


BUTTONS = (
    (
        "NewPatient",
        "Νέος ασθενής",
        "ToolbarNewPatient",
        "Προσθήκη νέου εσωτερικού ή εξωτερικού ασθενή.",
    ),
    (
        "NewProvider",
        "Θεραπευτής / Φοιτητής",
        "ToolbarNewProvider",
        "Προσθήκη νέου θεραπευτή ή φοιτητή.",
    ),
    (
        "PatientAbsence",
        "Απουσία ασθενή",
        "ToolbarPatientAbsence",
        "Καταχώρηση απουσίας ή ακύρωσης συνεδρίας ασθενή.",
    ),
    (
        "TherapistAbsence",
        "Απουσία θεραπευτή",
        "ToolbarTherapistAbsence",
        "Καταχώρηση απουσίας θεραπευτή και σχετικών ενεργειών.",
    ),
    (
        "TherapistDaily",
        "Ημερήσιο πρόγραμμα",
        "ToolbarTherapistDaily",
        "Μετάβαση στο ημερήσιο πρόγραμμα θεραπευτών.",
    ),
    (
        "Exit",
        "Save & Exit",
        "ExitApplication",
        "Αποθήκευση αλλαγών, κλείσιμο της εφαρμογής και επαναφορά του κανονικού Excel.",
    ),
)


class MasterToolbarError(RuntimeError):
    pass


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _find_master_header_row(ws) -> int:
    for row in range(1, 11):
        room = str(ws.Cells(row, 2).Value or "").strip().casefold()
        patient = str(ws.Cells(row, 3).Value or "").strip().casefold()
        if room == "θάλαμος".casefold() and patient == "ασθενής".casefold():
            return row
    raise MasterToolbarError("MASTER_SCHEDULE header row was not found")


def _delete_existing_toolbar_shapes(ws) -> None:
    names: list[str] = []
    for index in range(1, int(ws.Shapes.Count) + 1):
        shape = ws.Shapes(index)
        name = str(shape.Name)
        if name.startswith(MASTER_TOOLBAR_PREFIX):
            names.append(name)
    for name in names:
        ws.Shapes(name).Delete()


def _ensure_app_header_rows(ws) -> int:
    header_row = _find_master_header_row(ws)
    if header_row == 1:
        ws.Rows("1:3").Insert()
        header_row = 4
    if header_row != 4:
        raise MasterToolbarError(
            f"Unexpected MASTER_SCHEDULE header row {header_row}; expected 1 or 4"
        )
    ws.Rows(1).RowHeight = 6
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 6
    return header_row


def _ensure_master_display_schema(ws) -> None:
    header_row = _find_master_header_row(ws)

    col_j = str(ws.Cells(header_row, 10).Value or "").strip()
    col_k = str(ws.Cells(header_row, 11).Value or "").strip()
    col_l = str(ws.Cells(header_row, 12).Value or "").strip()
    col_m = str(ws.Cells(header_row, 13).Value or "").strip()

    if col_j.casefold() != "ΕΦΑ".casefold():
        raise MasterToolbarError("MASTER_SCHEDULE EFA column was not found in column J")

    if col_k.casefold() == "Κατάσταση".casefold():
        ws.Columns("K:L").Insert()
        ws.Cells(header_row, 11).Value = "ΨΥΧΟΛΟΓΟΙ"
        ws.Cells(header_row, 12).Value = "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ"
        ws.Cells(header_row, 13).Value = "Κατάσταση"
    elif (
        col_k.casefold() == "ΨΥΧΟΛΟΓΟΙ".casefold()
        and col_l.casefold() == "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ".casefold()
        and col_m.casefold() == "Κατάσταση".casefold()
    ):
        pass
    else:
        raise MasterToolbarError(
            "Unexpected MASTER_SCHEDULE columns after EFA; refusing to guess schema"
        )

    ws.Cells(header_row, 7).Value = "Ανακλ/μενο"
    ws.Columns("G").ColumnWidth = 11
    ws.Columns("G").WrapText = True
    ws.Columns("G").HorizontalAlignment = -4108
    ws.Columns("G").VerticalAlignment = -4108
    ws.Cells(header_row, 7).WrapText = True

    ws.Columns("K").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 20

    try:
        last_row = int(ws.Cells(ws.Rows.Count, 3).End(-4162).Row)
    except Exception:
        last_row = header_row
    last_row = max(header_row, last_row)

    data_range = ws.Range(f"K{header_row}:L{last_row}")
    data_range.WrapText = True

    # Light salmon for Psychology, light purple for Afternoon Program.
    ws.Range(f"K{header_row}:K{last_row}").Interior.Color = 13421823
    ws.Range(f"L{header_row}:L{last_row}").Interior.Color = 16764108

    for row in range(header_row + 1, last_row + 1):
        room = str(ws.Cells(row, 2).Value or "").strip()
        patient = str(ws.Cells(row, 3).Value or "").strip()
        if room and not patient:
            for col in range(13, 17):
                tail = ws.Cells(row, col).Value
                if tail is False or str(tail or "").strip().casefold() == "false":
                    ws.Cells(row, col).ClearContents()

    try:
        ws.ScrollArea = f"A1:M{last_row + 2}"
    except Exception:
        pass


def _freeze_master_identity_columns(ws, workbook) -> None:
    try:
        ws.Activate()
        window = workbook.Application.ActiveWindow
        window.FreezePanes = False
        window.SplitRow = 0
        window.SplitColumn = 3
        window.FreezePanes = True
    except Exception:
        pass


def _add_navigation_button(ws, *, name: str, caption: str, macro: str, left: float, top: float, width: float) -> None:
    try:
        ws.Shapes(name).Delete()
    except Exception:
        pass
    shape = ws.Shapes.AddShape(5, left, top, width, 24)
    shape.Name = name
    shape.OnAction = macro
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.Line.Weight = 1
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3


USER_FACING_SHEETS = ("DAILY_INPUT", "THERAPIST_DAILY", "REPLACEMENTS")


def _last_used_row(ws) -> int:
    try:
        used = ws.UsedRange
        return max(1, int(used.Row) + int(used.Rows.Count) - 1)
    except Exception:
        return 1


def _install_operational_navigation(workbook) -> None:
    for sheet_name in USER_FACING_SHEETS:
        try:
            ws = workbook.Worksheets(sheet_name)
        except Exception:
            continue

        last_row = _last_used_row(ws)
        anchor_row = last_row + 2
        try:
            anchor = ws.Range(f"A{anchor_row}")
            left = float(anchor.Left) + 4
            top = float(anchor.Top) + 2
        except Exception:
            left = 4.0
            top = float(anchor_row * 15)

        if sheet_name == "DAILY_INPUT":
            _add_navigation_button(
                ws,
                name=MASTER_TOOLBAR_PREFIX + "ApplyContinue",
                caption="Εφαρμογή & Συνέχεια",
                macro="ApplyDailyInputPreview",
                left=left,
                top=top,
                width=118,
            )
            nav_left = left + 124
        else:
            nav_left = left

        _add_navigation_button(
            ws,
            name=MASTER_TOOLBAR_PREFIX + "BackToMaster",
            caption="← MASTER",
            macro="GoToMaster",
            left=nav_left,
            top=top,
            width=82,
        )
        _add_navigation_button(
            ws,
            name=MASTER_TOOLBAR_PREFIX + "ExitApp",
            caption="Save & Exit",
            macro="ExitApplication",
            left=nav_left + 88,
            top=top,
            width=70,
        )


def install_master_toolbar(vbproject, workbook) -> None:
    _remove_component_if_present(vbproject, MASTER_TOOLBAR_MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MASTER_TOOLBAR_MODULE_NAME
    module.CodeModule.AddFromString(MASTER_TOOLBAR_MODULE_CODE)

    try:
        ws = workbook.Worksheets("MASTER_SCHEDULE")
    except Exception as exc:
        raise MasterToolbarError("Workbook has no MASTER_SCHEDULE sheet") from exc

    _ensure_app_header_rows(ws)
    _ensure_master_display_schema(ws)
    _delete_existing_toolbar_shapes(ws)
    _freeze_master_identity_columns(ws, workbook)

    left = float(ws.Range("A2").Left) + 4
    top = float(ws.Range("A2").Top) + 2
    available_width = float(ws.Range("A2:K2").Width) - 8
    gap = 6.0
    button_width = (available_width - gap * (len(BUTTONS) - 1)) / len(BUTTONS)
    button_height = max(24.0, float(ws.Rows(2).Height) - 4)

    for index, (key, caption, macro, description) in enumerate(BUTTONS):
        shape = ws.Shapes.AddShape(
            5,  # msoShapeRoundedRectangle
            left + index * (button_width + gap),
            top,
            button_width,
            button_height,
        )
        shape.Name = MASTER_TOOLBAR_PREFIX + key
        shape.OnAction = macro
        shape.AlternativeText = description
        shape.Placement = 3  # xlFreeFloating
        shape.Fill.ForeColor.RGB = 15527148  # RGB(236, 239, 244)
        shape.Line.ForeColor.RGB = 10066329  # RGB(153, 153, 153)
        shape.Line.Weight = 1
        shape.TextFrame2.TextRange.Text = caption
        shape.TextFrame2.TextRange.Font.Name = "Calibri"
        shape.TextFrame2.TextRange.Font.Size = 10
        shape.TextFrame2.TextRange.Font.Bold = True
        shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        shape.TextFrame2.VerticalAnchor = 3

    _install_operational_navigation(workbook)



def finalize_user_navigation(workbook_path) -> None:
    """Re-apply user-facing navigation after every build step has created its sheets."""
    from pathlib import Path
    import sys

    if sys.platform != "win32":
        raise MasterToolbarError("Navigation finalization requires Windows with Microsoft Excel")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise MasterToolbarError("pywin32 is required for navigation finalization") from exc

    path = Path(workbook_path).resolve()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        _install_operational_navigation(workbook)
        workbook.Save()
    except MasterToolbarError:
        raise
    except Exception as exc:
        raise MasterToolbarError(f"Could not finalize user navigation: {exc}") from exc
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
