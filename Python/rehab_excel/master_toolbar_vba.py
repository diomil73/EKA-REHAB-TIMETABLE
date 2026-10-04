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
    _delete_existing_toolbar_shapes(ws)

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
