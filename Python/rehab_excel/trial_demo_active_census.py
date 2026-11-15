"""Private native DEMO: live, source-linked patient status view.

This is an *additional* working-day census/exit view. The canonical MASTER
layout and immutable patient schedule are not altered. All visible values are
Excel formulas that refer to PATIENTS and PATIENT_EXITS, so confirmed exits
are reflected after Excel recalculates and after save/reopen.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys

from openpyxl.utils import get_column_letter

from .trial_demo_weekly_navigation import _shape


SHEET_NAME = "DEMO_ACTIVE_CENSUS"
MODULE_NAME = "modDemoActiveCensus"
OPEN_MACRO = "OpenDemoActiveCensus"
BUTTON_NAME = "diomilDemoActiveCensusOpen"
BACK_BUTTON = "diomilDemoActiveCensusBack"

HEADERS = (
    "PatientID", "Ασθενής", "Τύπος", "Θάλαμος",
    "Τελευταία ενεργή", "Κατάσταση",
)

VBA_CODE = r'''Option Explicit

' Extend the source-linked grid after registrations, without replacing the
' original patient registry or writing lifecycle events. Row 2 is a template.
Public Sub RefreshDemoActiveCensus()
    Dim patients As Worksheet, view As Worksheet
    Dim patientLast As Long, viewLast As Long
    Set patients = ThisWorkbook.Worksheets("PATIENTS")
    Set view = ThisWorkbook.Worksheets("DEMO_ACTIVE_CENSUS")
    patientLast = Application.Max(2, _
        patients.Cells(patients.Rows.Count, 1).End(xlUp).Row, _
        patients.Cells(patients.Rows.Count, 3).End(xlUp).Row)
    viewLast = Application.Max(2, _
        view.Cells(view.Rows.Count, 1).End(xlUp).Row)
    If Left$(CStr(view.Cells(2, 1).Formula), 4) <> "=IF(" Then
        Err.Raise vbObjectError + 3091, , "Η βάση της ενεργής λίστας λείπει."
    End If
    If patientLast > viewLast Then
        view.Range("A" & CStr(viewLast) & ":F" & CStr(viewLast)).AutoFill _
            Destination:=view.Range("A" & CStr(viewLast) & ":F" & CStr(patientLast)), _
            Type:=xlFillDefault
    End If
    view.Calculate
End Sub

Public Sub OpenDemoActiveCensus()
    On Error GoTo MissingSheet
    RefreshDemoActiveCensus
    With ThisWorkbook.Worksheets("DEMO_ACTIVE_CENSUS")
        .Visible = xlSheetVisible
        .Activate
        .Range("A1").Select
    End With
    KeepApplicationShell
    Exit Sub
MissingSheet:
    MsgBox "Η προβολή ενεργών ασθενών χρειάζεται έλεγχο.", vbExclamation, "Diomil DEMO"
End Sub
'''


class DemoActiveCensusError(ValueError):
    """Incomplete native demo status sheet or a missing navigation control."""


@dataclass(frozen=True)
class PatientCensusFormulas:
    identity: str
    name: str
    patient_type: str
    room: str
    last_active: str
    status: str


def census_formulas(
    row: int, *,
    patient_type_source_column: int,
) -> PatientCensusFormulas:
    """Return exact Excel formulas, without accessing any private identities."""
    if row < 2 or patient_type_source_column < 6:
        raise DemoActiveCensusError("Invalid patient census row/type column")
    patient_type_col = get_column_letter(patient_type_source_column)
    base = f'PATIENTS!C{row}=""'
    return PatientCensusFormulas(
        identity=f'=IF({base},"",PATIENTS!A{row})',
        name=f'=IF({base},"",PATIENTS!C{row})',
        patient_type=f'=IF({base},"",PATIENTS!{patient_type_col}{row})',
        room=f'=IF({base},"",PATIENTS!B{row})',
        # Full-column lookup does not silently stop seeing events beyond an
        # arbitrary fixed row count as the practice registry grows.
        last_active=(
            f'=IF(A{row}="","",IFERROR('
            f'INDEX(PATIENT_EXITS!$C:$C,'
            f'MATCH(A{row},PATIENT_EXITS!$A:$A,0)),""))'
        ),
        status=(
            f'=IF(A{row}="","",IF(E{row}="","Ενεργός",'
            f'IF(INT(E{row})>=INT($H$2),"Ενεργός","Εξιτήριο / Λήξη")))'
        ),
    )


def install_demo_active_census(workbook, *, patient_rows: int) -> None:
    """Install recalculating census, and independent MASTER/back navigation."""
    if patient_rows < 0:
        raise DemoActiveCensusError("Invalid patient row count")
    master = workbook.Worksheets("MASTER_SCHEDULE")
    patients = workbook.Worksheets("PATIENTS")
    try:
        ws = workbook.Worksheets(SHEET_NAME)
    except Exception:
        ws = workbook.Worksheets.Add(
            After=workbook.Worksheets(workbook.Worksheets.Count)
        )
        ws.Name = SHEET_NAME

    headers = {}
    for col in range(1, max(9, int(patients.UsedRange.Columns.Count)) + 1):
        text = str(patients.Cells(1, col).Value2 or "").strip()
        if text:
            headers[text] = col
    type_col = headers.get("PatientType")
    if type_col is None:
        raise DemoActiveCensusError("PATIENTS PatientType column is required")
    if int(type_col) < 6:
        raise DemoActiveCensusError("Unsafe PatientType column position")
    for col, label in enumerate(HEADERS, start=1):
        existing = str(ws.Cells(1, col).Value2 or "").strip()
        if existing not in ("", label):
            raise DemoActiveCensusError("Existing active census schema mismatch")
        ws.Cells(1, col).Value = label
    ws.Cells(1, 8).Value = "Ημερομηνία ελέγχου"
    if ws.Cells(2, 8).Value2 in (None, ""):
        ws.Cells(2, 8).Formula = "=TODAY()"
    ws.Cells(2, 8).NumberFormat = "dd/mm/yyyy"
    # Row 2 remains a blank, formula-only template even if the DEMO starts
    # with zero patients. VBA can extend the grid after future admissions.
    for row in range(2, max(3, patient_rows + 2)):
        formula = census_formulas(row, patient_type_source_column=int(type_col))
        for col, value in enumerate(
            (
                formula.identity, formula.name, formula.patient_type,
                formula.room, formula.last_active, formula.status,
            ), start=1,
        ):
            if ws.Cells(row, col).Value2 not in (None, ""):
                raise DemoActiveCensusError("Active census row is occupied")
            ws.Cells(row, col).Formula = value
        ws.Cells(row, 5).NumberFormat = "dd/mm/yyyy"
    ws.Range("A1:F1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("H").ColumnWidth = 22

    components = workbook.VBProject.VBComponents
    try:
        module = components(MODULE_NAME)
        n = int(module.CodeModule.CountOfLines)
        if n:
            module.CodeModule.DeleteLines(1, n)
    except Exception:
        module = components.Add(1)
        module.Name = MODULE_NAME
    module.CodeModule.AddFromString(VBA_CODE)
    origin = master.Range("A2")
    _shape(
        master, name=BUTTON_NAME, caption="Ενεργοί ασθενείς",
        action=OPEN_MACRO,
        left=float(origin.Left) + 716, top=float(origin.Top) + 5,
        width=128,
    )
    back = ws.Range("I1")
    _shape(
        ws, name=BACK_BUTTON, caption="← MASTER", action="GoToMaster",
        left=float(back.Left) + 4, top=float(back.Top) + 2,
        width=90,
    )


def verify_demo_active_census(path: str | Path, *, patient_rows: int) -> None:
    """Native Excel read-only checks of saved VBA and the entire formula grid."""
    if sys.platform != "win32":
        raise DemoActiveCensusError("Native Microsoft Excel for Windows required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        raise DemoActiveCensusError("pywin32 required") from None

    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True)
        module = wb.VBProject.VBComponents(MODULE_NAME)
        code = str(module.CodeModule.Lines(
            1, int(module.CodeModule.CountOfLines)
        ))
        if ("Public Sub OpenDemoActiveCensus()" not in code
                or "Public Sub RefreshDemoActiveCensus()" not in code):
            raise DemoActiveCensusError("Active census VBA module missing")
        master = wb.Worksheets("MASTER_SCHEDULE")
        ws = wb.Worksheets(SHEET_NAME)
        if (
            str(master.Shapes(BUTTON_NAME).OnAction).split("!")[-1] != OPEN_MACRO
            or str(ws.Shapes(BACK_BUTTON).OnAction).split("!")[-1] != "GoToMaster"
        ):
            raise DemoActiveCensusError("Active census navigation mismatch")
        if tuple(str(ws.Cells(1, i).Value2 or "") for i in range(1, 7)) != HEADERS:
            raise DemoActiveCensusError("Active census headers missing")
        if str(ws.Cells(2, 8).Formula or "").upper() != "=TODAY()":
            raise DemoActiveCensusError("Active census date formula missing")
        patients = wb.Worksheets("PATIENTS")
        type_col = next((
            col for col in range(6, int(patients.UsedRange.Columns.Count) + 1)
            if str(patients.Cells(1, col).Value2 or "").strip() == "PatientType"
        ), None)
        if type_col is None:
            raise DemoActiveCensusError("Patient type is not available")
        for row in range(2, max(3, patient_rows + 2)):
            expected = census_formulas(
                row, patient_type_source_column=type_col,
            )
            for col, text in enumerate((
                expected.identity, expected.name, expected.patient_type,
                expected.room, expected.last_active, expected.status,
            ), start=1):
                if str(ws.Cells(row, col).Formula or "") != text:
                    raise DemoActiveCensusError(
                        "Saved active census formula differs from source"
                    )
    except DemoActiveCensusError:
        raise
    except Exception:
        raise DemoActiveCensusError(
            "SAFETY STOP: active census native readback failed"
        ) from None
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
