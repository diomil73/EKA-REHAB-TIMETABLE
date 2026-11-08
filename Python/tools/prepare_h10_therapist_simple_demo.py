from __future__ import annotations

"""One-shot native H10 therapist UI refinement on a fresh copy of H8 BO.

Copies already verified synthetic history intact. Does not rebuild seed data,
rewrite the snapshot stores or modify BO. All native edits take place in a
previously nonexistent candidate only.
"""

import argparse
import hashlib
from pathlib import Path
import shutil
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.therapist_history_report_vba import (
    THERAPIST_HISTORY_MODULE_NAME,
    THERAPIST_HISTORY_MODULE_CODE,
    THERAPIST_HISTORY_SHEET_NAME,
)

EVENT_CODE = """Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target.CountLarge <> 1 Then Exit Sub
    If Target.Row < 23 Then Exit Sub
    If Len(Trim$(CStr(Me.Cells(Target.Row, 3).Value2))) = 0 Then Exit Sub
    ShowSelectedTherapistChart
SafeExit:
End Sub
"""

REQUIRED_VBA = (
    'ActiveWindow.SplitRow = 22',
    'Formula1:="=ROW()=$P$1"',
    '"Μ.Ο. περιόδου: "',
    '"Απουσίες/αναβολές"',
    "Public Sub HighlightSelectedTherapistRow",
)


class H10DemoError(RuntimeError):
    pass


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_candidate(source: Path, output: Path) -> None:
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise H10DemoError("Existing verified BO .xlsm source required")
    if "_DEMO_" not in source.name.upper() or not source.name.upper().endswith("_BO.XLSM"):
        raise H10DemoError("Source must be frozen verified BO demo")
    if output.exists():
        raise H10DemoError("Candidate already exists; never overwrite")
    if "_DEMO_" not in output.name.upper() or not output.name.upper().endswith("_BQ.XLSM"):
        raise H10DemoError("Fresh H10 BQ _DEMO_ .xlsm destination required")
    if source.resolve() == output.resolve():
        raise H10DemoError("Source and destination collision")


def _safe_module_lines(module) -> str:
    """Do not call VBA CodeModule.Lines(1, 0) on an empty worksheet module."""
    count = int(module.CountOfLines)
    return "" if count == 0 else str(module.Lines(1, count))


def _verify_native(wb) -> None:
    if not bool(wb.HasVBProject):
        raise H10DemoError("VBA project missing")
    module = wb.VBProject.VBComponents(THERAPIST_HISTORY_MODULE_NAME).CodeModule
    installed = _safe_module_lines(module)
    if any(fragment not in installed for fragment in REQUIRED_VBA):
        raise H10DemoError("H10 VBA markers missing after native reopen")
    sheet = wb.Worksheets(THERAPIST_HISTORY_SHEET_NAME)
    component = wb.VBProject.VBComponents(sheet.CodeName).CodeModule
    sheet_code = _safe_module_lines(component)
    normalized_sheet = sheet_code.replace("\r\n", "\n").replace("\r", "\n").strip()
    normalized_event = EVENT_CODE.replace("\r\n", "\n").replace("\r", "\n").strip()
    if normalized_event not in normalized_sheet:
        raise H10DemoError("Therapist row-selection event not persisted")
    if not bool(sheet.ProtectContents) or bool(sheet.ProtectDrawingObjects):
        raise H10DemoError("Therapist report protection invalid")
    for name in ("FINALIZED_DAY_SNAPSHOTS", "THERAPIST_DAY_HISTORY",
                 "PATIENT_SESSION_HISTORY", "DAY_HISTORY_EVENTS",
                 "THERAPIST_ATTENDANCE_HISTORY"):
        technical = wb.Worksheets(name)
        if int(technical.Visible) != 2 or not bool(technical.ProtectContents):
            raise H10DemoError(f"Frozen technical history altered/unprotected: {name}")


def install_native(source: Path, output: Path) -> None:
    check_candidate(source, output)
    if sys.platform != "win32":
        raise H10DemoError("Native H10 installation requires Windows and Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise H10DemoError("pywin32 required") from exc

    source_hash = digest(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    with source.open("rb") as inp, output.open("xb") as out:
        shutil.copyfileobj(inp, out)
    if digest(output) != source_hash or digest(source) != source_hash:
        raise H10DemoError("Fresh BO -> BQ copy SHA256 mismatch")

    excel = None
    book = None
    stage = "initialize native Excel session"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        stage = "open BQ for installation"
        book = excel.Workbooks.Open(str(output.resolve()), UpdateLinks=0, ReadOnly=False)
        if bool(book.ReadOnly):
            raise H10DemoError("BQ opened read-only; no edits possible")
        vbproj = book.VBProject
        report = book.Worksheets(THERAPIST_HISTORY_SHEET_NAME)
        stage = "locate therapist report sheet module"
        component = vbproj.VBComponents(report.CodeName)
        code = component.CodeModule
        stage = "read existing worksheet VBA (may be empty)"
        present = _safe_module_lines(code)
        if "Worksheet_SelectionChange" in present:
            raise H10DemoError("Existing selection event detected; stop rather than overwrite")
        stage = "replace therapist history standard module"
        old = vbproj.VBComponents(THERAPIST_HISTORY_MODULE_NAME)
        vbproj.VBComponents.Remove(old)
        new = vbproj.VBComponents.Add(1)
        new.Name = THERAPIST_HISTORY_MODULE_NAME
        stage = "write therapist history standard VBA"
        new.CodeModule.AddFromString(THERAPIST_HISTORY_MODULE_CODE)
        stage = "write worksheet selection-change handler"
        code.AddFromString(EVENT_CODE)
        stage = "verify installed VBA and protected history sheets"
        _verify_native(book)
        stage = "save BQ"
        book.Save()
    except Exception as exc:
        raise H10DemoError(f"BQ install failed at {stage}: {exc}") from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            excel.Quit()

    excel = None
    book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        stage = "read-only persisted BQ reopen"
        book = excel.Workbooks.Open(str(output.resolve()), UpdateLinks=0, ReadOnly=True)
        if not bool(book.ReadOnly):
            raise H10DemoError("Native BQ reopen was not read-only")
        stage = "read-only persisted VBA protection audit"
        _verify_native(book)
    except Exception as exc:
        raise H10DemoError(f"BQ verification failed at {stage}: {exc}") from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            excel.Quit()

    if digest(source) != source_hash:
        raise H10DemoError("Frozen BO source hash changed")
    print("H10 THERAPIST SIMPLE DEMO INSTALL OK")
    print("Fresh BQ only; BO and archived previews unchanged")
    print("Sticky therapist charts, selection highlight, period benchmarks installed")
    print("Synthetic data is demonstration only; no production attendance KPI")


def main() -> int:
    parser = argparse.ArgumentParser(description="Install H10 simple therapist visuals on a fresh BO copy")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        install_native(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        print("Freeze BQ; do not retry, repair or overwrite candidate")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
