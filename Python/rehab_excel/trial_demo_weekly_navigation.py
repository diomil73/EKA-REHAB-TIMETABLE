"""Demo-only MASTER <-> full weekly appointments navigation.

The existing MASTER toolbar has a compact, fixed four-button contract. This
module adds one separately named practice button without replacing that toolbar,
and a return button on the detailed weekly view. It is installed only into a
fresh isolated native Excel DEMO stage, never the operational source or a frozen
proof workbook.
"""
from __future__ import annotations

from pathlib import Path
import sys


DEMO_NAV_MODULE = "modDemoWeeklyNavigation"
MASTER_BUTTON = "diomilDemoWeeklyOpen"
BACK_BUTTON = "diomilDemoWeeklyBack"
VIEW_SHEET = "DEMO_WEEKLY_VIEW"
OPEN_MACRO = "OpenDemoWeeklyView"
BACK_MACRO = "GoToMaster"

VBA_CODE = r'''Option Explicit

Public Sub OpenDemoWeeklyView()
    On Error GoTo MissingView
    ThisWorkbook.Worksheets("DEMO_WEEKLY_VIEW").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("DEMO_WEEKLY_VIEW").Activate
    ThisWorkbook.Worksheets("DEMO_WEEKLY_VIEW").Range("A1").Select
    KeepApplicationShell
    Exit Sub
MissingView:
    MsgBox "Δεν βρέθηκε το πλήρες εβδομαδιαίο πρόγραμμα.", vbExclamation, "Diomil DEMO"
End Sub
'''


class DemoWeeklyNavigationError(RuntimeError):
    pass


def _shape(ws, *, name: str, caption: str, action: str, left: float,
           top: float, width: float) -> None:
    try:
        ws.Shapes(name).Delete()
    except Exception:
        pass
    shape = ws.Shapes.AddShape(5, left, top, width, 24)
    shape.Name = name
    shape.OnAction = action
    shape.Placement = 3
    shape.AlternativeText = "DIOMIL practice DEMO navigation"
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329


def install_demo_weekly_navigation(workbook) -> None:
    """Install navigation after the existing full application VBA contract."""
    master = workbook.Worksheets("MASTER_SCHEDULE")
    view = workbook.Worksheets(VIEW_SHEET)
    vbproject = workbook.VBProject
    try:
        module = vbproject.VBComponents(DEMO_NAV_MODULE)
        module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines)
    except Exception:
        module = vbproject.VBComponents.Add(1)
        module.Name = DEMO_NAV_MODULE
    module.CodeModule.AddFromString(VBA_CODE)

    origin = master.Range("A2")
    # The four canonical toolbar buttons occupy roughly the first 430 px.
    _shape(
        master, name=MASTER_BUTTON, caption="Πλήρες εβδομαδιαίο",
        action=OPEN_MACRO,
        left=float(origin.Left) + 444, top=float(origin.Top) + 5,
        width=128,
    )
    anchor = view.Range("I1")
    _shape(
        view, name=BACK_BUTTON, caption="← MASTER",
        action=BACK_MACRO,
        left=float(anchor.Left) + 4, top=float(anchor.Top) + 2,
        width=90,
    )


def verify_demo_weekly_navigation(workbook_path: str | Path) -> None:
    """Read-only native Excel check of both actual buttons and macro body."""
    if sys.platform != "win32":
        raise DemoWeeklyNavigationError("Native Excel is required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        raise DemoWeeklyNavigationError("pywin32 is required") from None
    excel = None
    book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        book = excel.Workbooks.Open(
            str(Path(workbook_path).resolve()), UpdateLinks=0, ReadOnly=True,
        )
        module = book.VBProject.VBComponents(DEMO_NAV_MODULE)
        code = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)
        if "Public Sub OpenDemoWeeklyView()" not in code:
            raise DemoWeeklyNavigationError("Missing DEMO view macro")
        master = book.Worksheets("MASTER_SCHEDULE")
        view = book.Worksheets(VIEW_SHEET)
        if (
            str(master.Shapes(MASTER_BUTTON).OnAction).split("!")[-1] != OPEN_MACRO
            or str(view.Shapes(BACK_BUTTON).OnAction).split("!")[-1] != BACK_MACRO
        ):
            raise DemoWeeklyNavigationError("Practice navigation button action differs")
    except DemoWeeklyNavigationError:
        raise
    except Exception:
        raise DemoWeeklyNavigationError(
            "SAFETY STOP: native DEMO navigation was not saved correctly"
        ) from None
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
