from __future__ import annotations

"""One-shot native H10 therapist UI refinement on fresh isolated H8 CB.

Copies independently verified synthetic-only CB history intact. Does not rebuild seed data,
rewrite snapshot stores or modify CA/CB. All native edits take place in a
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

from tools.verify_h8_synthetic_history import verify_demo_history
from tools.report_therapist_history import report_therapist_history
from tools.audit_therapist_daily_capacity import capacity_anomalies
from tools.audit_h10_baseline_clock_conflicts import clock_conflicts
from rehab_excel.patient_session_history_store import read_current_patient_session_history_rows
from rehab_excel.therapist_history_report_vba import (
    THERAPIST_HISTORY_MODULE_NAME,
    THERAPIST_HISTORY_MODULE_CODE,
    THERAPIST_HISTORY_SHEET_NAME,
)

EVENT_CODE = """Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target.CountLarge <> 1 Then Exit Sub
    If Target.Row < 23 Then Exit Sub
    ' The aggregate appendix is below the therapist table; never treat it
    ' as an individual therapist even though it has numeric column C.
    If Target.Row > CLng(Me.Range("P3").Value2) Then Exit Sub
    If Len(Trim$(CStr(Me.Cells(Target.Row, 3).Value2))) = 0 Then Exit Sub
    ShowSelectedTherapistChart
SafeExit:
End Sub
"""

REQUIRED_VBA = (
    'ActiveWindow.SplitRow = 22',
    'Private Sub AddTherapistMeanMarker',
    'plot.Chart.Shapes.AddLine',
    'RGB(57, 255, 20)',
    '" | Υλοποίηση: "',
    'cohortMaxSessions = Application.Max',
    'plot.Chart.Axes(xlValue).MaximumScale = cohortAxisSessions',
    'Private Sub BuildTherapistMiniTrends',
    'BuildTherapistMiniTrends ws, reportMode, selectedRow, therapistName',
    'AddTherapistTrendSeries plot, "Αρχικά προγραμματισμένες"',
    'AddTherapistTrendSeries plot, "Τελικά πραγματοποιημένες"',
    'ws.Columns("P:U").Hidden = True',
    'Private Sub RenderManagementSummary',
    'ws.Range("A8:B8").Merge',
    'ws.Range("F15").Top',
    'shape.Top = ws.Range("A5").Top',
    'ws.Range("A22:N2000").ClearFormats',
    'ws.Cells(lastGroup + 1, 3)).NumberFormat = "0"',
    'Public Sub ExportManagementOnlyWorkbook',
    'Public Sub ShowManagementSummaryDaily',
    'Public Sub ShowManagementSummaryOverall',
    'Private Sub BuildDepartmentAppendix',
    'chart.Name = "ekaDepartmentHistoryChart"',
    'ws.Range("P3").Value2 = rowIndex - 1',
    'If positiveDeltaDays > 0 Then',
    'HasMajorGridlines = True',
    'TRIAL #FF0303',
    'TRIAL #5BFA74',
    'Formula1:="=ROW()=$P$1"',
    '"Μ.Ο.: "',
    'cohortAxisPatients = 6',
    'If loadPerDay > 6 Then',
    'plot.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 9',
    '"Απουσίες/αναβολές"',
    "Public Sub HighlightSelectedTherapistRow",
)


class H10DemoError(RuntimeError):
    pass


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_candidate(source: Path, output: Path) -> None:
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise H10DemoError("Existing synthetic-only CB .xlsm source required")
    if "_DEMO_" not in source.name.upper() or not source.name.upper().endswith("_CB.XLSM"):
        raise H10DemoError("Source must be isolated CB demo, never CA/BO")
    if output.exists():
        raise H10DemoError("Candidate already exists; never overwrite")
    if "_DEMO_" not in output.name.upper() or not output.name.upper().endswith("_BZ.XLSM"):
        raise H10DemoError("Fresh H10 BZ _DEMO_ .xlsm destination required")
    if source.resolve() == output.resolve():
        raise H10DemoError("Source and destination collision")


def _safe_module_lines(module) -> str:
    """Do not call VBA CodeModule.Lines(1, 0) on an empty worksheet module."""
    count = int(module.CountOfLines)
    return "" if count == 0 else str(module.Lines(1, count))


def _event_lines(source: str) -> list[str]:
    """Compare VBA statements, not VBIDE's whitespace/line-ending presentation."""
    return [
        " ".join(line.strip().casefold().split())
        for line in source.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if line.strip()
    ]


def verify_selection_event(source: str) -> None:
    """Fail closed on missing/duplicate event or a missing safety guard.

    VBIDE may reformat spaces or line endings on CodeModule.AddFromString.
    Literal whole-block comparison is not a sound persisted-code contract.
    """
    actual = _event_lines(source)
    required = _event_lines(EVENT_CODE)
    signature = required[0]
    if actual.count(signature) != 1:
        raise H10DemoError(
            f"Therapist selection event declaration count {actual.count(signature)}; "
            f"VBA lines={len(actual)}"
        )
    offset = actual.index(signature)
    if actual[offset:offset + len(required)] != required:
        # Report only concise diagnostic metadata, never assume installation.
        mismatch = next(
            (i for i, line in enumerate(required)
             if offset + i >= len(actual) or actual[offset + i] != line),
            len(required),
        )
        raise H10DemoError(
            f"Therapist selection event failed statement readback at "
            f"relative line {mismatch + 1}; VBA lines={len(actual)}"
        )


def _verify_native(wb) -> None:
    if not bool(wb.HasVBProject):
        raise H10DemoError("VBA project missing")
    module = wb.VBProject.VBComponents(THERAPIST_HISTORY_MODULE_NAME).CodeModule
    installed = _safe_module_lines(module)
    # VBIDE may normalize whitespace during import. Keep the same persisted
    # contract but identify which marker was not read back.
    normalize = lambda value: " ".join(value.split()).casefold()
    installed_normalized = normalize(installed)
    missing = [
        fragment for fragment in REQUIRED_VBA
        if normalize(fragment) not in installed_normalized
    ]
    if missing:
        raise H10DemoError(
            "H10 VBA marker read-back mismatch: " + repr(missing)
            + f"; persisted lines={int(module.CountOfLines)}"
        )
    sheet = wb.Worksheets(THERAPIST_HISTORY_SHEET_NAME)
    management = wb.Worksheets("MANAGEMENT_SUMMARY")
    managed_contents = bool(management.ProtectContents)
    managed_drawing = bool(management.ProtectDrawingObjects)
    managed_scenarios = bool(management.ProtectScenarios)
    if not managed_contents or managed_drawing or not managed_scenarios:
        raise H10DemoError(
            "Management preview protection invalid: "
            f"Contents={managed_contents}; DrawingObjects={managed_drawing}; "
            f"Scenarios={managed_scenarios}"
        )
    component = wb.VBProject.VBComponents(sheet.CodeName).CodeModule
    sheet_code = _safe_module_lines(component)
    verify_selection_event(sheet_code)
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

    # Fail closed BEFORE creating the destination. The old frozen CA has
    # known synthetic 8/9 occupied-slot errors and must never feed BZ.
    verify_demo_history(source)
    # The CB admission invariant is stronger than H8 synthetic correctness:
    # CB must contain NO inherited test-archive rows in any history ledger.
    from tools.prepare_h10_synthetic_only_source import classify_rows
    for name, (_, non_demo) in classify_rows(source).items():
        if non_demo:
            raise H10DemoError(f"CB contains archived test history in {name}")
    anomalies = capacity_anomalies(report_therapist_history(source), {})
    blocking = [
        row for row in anomalies if
        "OCCUPIED_SLOTS_EXCEED_CAPACITY" in row["flags"]
        or "CAPACITY_NOT_CONFIGURED" in row["flags"]
    ]
    if blocking:
        raise H10DemoError(
            f"Isolated CB capacity gate failed: {len(blocking)} therapist-days"
        )
    # A capacity-only gate cannot catch two patients occupying the same
    # clock slot with zero recorded approval (legacy Oct 30 evidence).
    # Inspect ALL finalized service dates and stop BEFORE opening output.
    reports = report_therapist_history(source)
    sessions = read_current_patient_session_history_rows(source)
    service_dates = sorted({row.service_date for row in reports})
    clock_flags = [
        exception
        for service_date in service_dates
        for exception in clock_conflicts(reports, sessions, service_date)
        if "SHARED_CLOCK_WITHOUT_RECORDED_APPROVAL" in exception["flags"]
        or "SLOT_OR_SESSION_RECONCILIATION_MISMATCH" in exception["flags"]
    ]
    if clock_flags:
        raise H10DemoError(
            f"Isolated CB clock/approval gate failed: {len(clock_flags)} therapist-days"
        )
    source_hash = digest(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    with source.open("rb") as inp, output.open("xb") as out:
        shutil.copyfileobj(inp, out)
    if digest(output) != source_hash or digest(source) != source_hash:
        raise H10DemoError("Fresh CB -> BZ copy SHA256 mismatch")

    excel = None
    book = None
    stage = "initialize native Excel session"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        stage = "open BZ for installation"
        book = excel.Workbooks.Open(str(output.resolve()), UpdateLinks=0, ReadOnly=False)
        if bool(book.ReadOnly):
            raise H10DemoError("BZ opened read-only; no edits possible")
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
        stage = "verify freshly written standard VBA"
        imported = _safe_module_lines(new.CodeModule)
        if not imported:
            raise H10DemoError("Imported standard VBA unexpectedly empty")
        stage = "write worksheet selection-change handler"
        code.AddFromString(EVENT_CODE)
        stage = "install protected management-only preview sheet"
        try:
            management = book.Worksheets("MANAGEMENT_SUMMARY")
        except Exception:
            management = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
            management.Name = "MANAGEMENT_SUMMARY"
        management.Unprotect("")
        management.Cells.ClearContents()
        management.Range("A1").Value2 = (
            "ΔΙΟΙΚΗΣΗ | Ανοίξτε «Προβολή διοίκησης» από τα στατιστικά θεραπευτών"
        )
        management.Range("A2").Value2 = (
            "Μην αποστείλετε το λειτουργικό XLSM. Χρησιμοποιήστε την ασφαλή εξαγωγή XLSX."
        )
        # Explicit POSITIONAL COM arguments: Excel Worksheet.Protect
        # signature is (Password, DrawingObjects, Contents, Scenarios, ...).
        # Keep drawings editable for dashboard buttons and chart refresh.
        # The readback strictly verifies actual Excel protection flags.
        management.Protect("", False, True, True)
        stage = "verify installed VBA and protected history sheets"
        _verify_native(book)
        stage = "save BZ"
        book.Save()
    except Exception as exc:
        raise H10DemoError(f"BZ install failed at {stage}: {exc}") from exc
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
        stage = "read-only persisted BZ reopen"
        book = excel.Workbooks.Open(str(output.resolve()), UpdateLinks=0, ReadOnly=True)
        if not bool(book.ReadOnly):
            raise H10DemoError("Native BZ reopen was not read-only")
        stage = "read-only persisted VBA protection audit"
        _verify_native(book)
    except Exception as exc:
        raise H10DemoError(f"BZ verification failed at {stage}: {exc}") from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            excel.Quit()

    if digest(source) != source_hash:
        raise H10DemoError("Isolated CB source hash changed")
    print("H10 THERAPIST SIMPLE DEMO INSTALL OK")
    print("Fresh BZ only; isolated CB and archived previews unchanged")
    print("Sticky therapist charts, selection highlight, period benchmarks installed")
    print("Synthetic data is demonstration only; no production attendance KPI")


def main() -> int:
    parser = argparse.ArgumentParser(description="Install H10 therapist visuals on a fresh synthetic-only CB copy")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        install_native(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        print("Freeze BZ; do not retry, repair or overwrite candidate")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
