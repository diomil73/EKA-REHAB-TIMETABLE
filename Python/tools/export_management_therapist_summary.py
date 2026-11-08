from __future__ import annotations

"""Export a strictly whitelisted, standalone, macro-free department report.

This script NEVER copies worksheets out of the operational .xlsm. It creates
a blank workbook using Excel on Windows and writes only total counts grouped
by date/week/month/whole history. Four friendly navigation tabs act as the
mode buttons in a file that can safely be shared without VBA/hidden sheets.
"""

import argparse
import hashlib
import math
from collections import defaultdict
from datetime import timedelta
from pathlib import Path
import sys
import zipfile

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.report_therapist_history import report_therapist_history


MODES = ("daily", "weekly", "monthly", "overall")
MODE_LABELS = {
    "daily": "ΗΜΕΡΗΣΙΑ",
    "weekly": "ΕΒΔΟΜΑΔΙΑΙΑ",
    "monthly": "ΜΗΝΙΑΙΑ",
    "overall": "ΣΥΝΟΛΙΚΑ",
}
BLUE_RGB = 0xF5A83F  # VBA/Excel BGR: #3FA8F5
RED_RGB = 0x0303FF   # VBA/Excel BGR: #FF0303


def department_periods(reports) -> dict[str, list[tuple[str, int, int, float | None]]]:
    """Only aggregate numbers; no therapist IDs, patient IDs or row pointers."""
    by_mode = {mode: defaultdict(lambda: [0, 0]) for mode in MODES}
    observed = set()
    for record in reports:
        key = (record.service_date, record.therapist_id.casefold())
        if key in observed:
            raise ValueError("Duplicate finalized therapist/day report")
        observed.add(key)
        planned = int(record.planned_sessions)
        actual = int(record.final_active_sessions)
        if min(planned, actual) < 0:
            raise ValueError("Negative finalized session count")
        day = record.service_date
        labels = {
            "daily": day.isoformat(),
            "weekly": (day - timedelta(days=day.weekday())).isoformat(),
            "monthly": day.strftime("%Y-%m"),
            "overall": "ΟΛΟ ΤΟ ΙΣΤΟΡΙΚΟ",
        }
        for mode, label in labels.items():
            by_mode[mode][label][0] += planned
            by_mode[mode][label][1] += actual

    return {
        mode: [
            (label, counts[0], counts[1],
             counts[1] / counts[0] if counts[0] > 0 else None)
            for label, counts in sorted(by_mode[mode].items())
        ]
        for mode in MODES
    }


def inspect_share_package(file: Path) -> None:
    """Fail closed if the ZIP contains macros, external links or extra sheets."""
    with zipfile.ZipFile(file) as zipped:
        files = set(zipped.namelist())
        suspicious = ("externalLinks/", "embeddings/", "connections.xml",
                      "queryTables/", "pivotCache", "vbaProject", "customXml/",
                      "activeX/", "oleObjects/")
        if any(any(token in name for token in suspicious) for name in files):
            raise ValueError("Unsafe VBA or external link or embedded data in management export")
        sheets = [
            name for name in files
            if name.startswith("xl/worksheets/sheet")
            and name.endswith(".xml") and "/" not in name[len("xl/worksheets/"):]
        ]
        if len(sheets) != 4:
            raise ValueError("Management export must contain exactly four standalone sheets")
        forbidden = (
            "FINALIZED_DAY_SNAPSHOTS", "THERAPIST_DAY_HISTORY",
            "PATIENT_SESSION_HISTORY", "THERAPIST_ATTENDANCE_HISTORY",
            "DAY_HISTORY_EVENTS", "modTherapistHistoryReport",
        )
        for name in files:
            if not name.endswith((".xml", ".rels")):
                continue
            content = zipped.read(name)
            if any(marker.encode() in content for marker in forbidden):
                raise ValueError("Operational schema leaked into management report")
            if name.endswith(".rels") and b'TargetMode="External"' in content:
                raise ValueError("External relationship in management report")
            if name == "xl/workbook.xml":
                if b'state="hidden"' in content or b'state="veryHidden"' in content:
                    raise ValueError("Hidden worksheet in management report")


def _write_mode_sheet(ws, mode, rows):
    ws.Name = MODE_LABELS[mode]
    ws.Cells.Font.Name = "Calibri"
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B:C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E:J").ColumnWidth = 13

    ws.Range("A1:J2").Merge()
    ws.Range("A1").Value2 = "ΕΙΚΟΝΑ ΤΜΗΜΑΤΟΣ | ΟΛΟΙ ΟΙ ΘΕΡΑΠΕΥΤΕΣ"
    ws.Range("A1:J2").Interior.Color = 0x3D2310
    ws.Range("A1:J2").Font.Color = 0xFFFFFF
    ws.Range("A1:J2").Font.Bold = True
    ws.Range("A1:J2").Font.Size = 17
    ws.Rows("1:2").RowHeight = 24

    for i, (tab, name) in enumerate(MODE_LABELS.items()):
        cell = ws.Cells(4, 1 + i * 2)
        cell.Value2 = name
        ws.Hyperlinks.Add(Anchor=cell, Address="", SubAddress=f"'{name}'!A1",
                          TextToDisplay=name)
        cell.Font.Bold = True
    ws.Range("A6:D6").Value2 = [[
        "Περίοδος", "Αρχικά προγραμματισμένες",
        "Τελικά πραγματοποιημένες", "Υλοποίηση",
    ]]
    ws.Range("A6:D6").Font.Bold = True
    ws.Range("A6:D6").Interior.Color = 0xE8DDE1
    planned_total = delivered_total = 0
    for i, (period, planned, delivered, pct) in enumerate(rows, 7):
        ws.Cells(i, 1).NumberFormat = "@"
        ws.Cells(i, 1).Value2 = period
        ws.Cells(i, 2).Value2 = planned
        ws.Cells(i, 3).Value2 = delivered
        if pct is None:
            ws.Cells(i, 4).Value2 = "Μη διαθέσιμο"
        else:
            ws.Cells(i, 4).Value2 = pct
            ws.Cells(i, 4).NumberFormat = "0.0%"
        planned_total += planned
        delivered_total += delivered
    final = 7 + len(rows)
    ws.Cells(final, 1).Value2 = "ΣΥΝΟΛΟ ΟΜΑΔΑΣ"
    ws.Cells(final, 2).Value2 = planned_total
    ws.Cells(final, 3).Value2 = delivered_total
    if planned_total > 0:
        ws.Cells(final, 4).Value2 = delivered_total / planned_total
        ws.Cells(final, 4).NumberFormat = "0.0%"
    else:
        ws.Cells(final, 4).Value2 = "Μη διαθέσιμο"
    ws.Range(ws.Cells(final, 1), ws.Cells(final, 4)).Font.Bold = True
    ws.Cells(final + 2, 1).Value2 = (
        "Μπλε: αρχικό πρόγραμμα | Κόκκινο: τελικό αποτέλεσμα. "
        "Μερικές ημέρες οι επιπλέον θεραπείες μπορεί να υπερβούν το αρχικό πρόγραμμα."
    )
    ws.Range(ws.Cells(final + 2, 1), ws.Cells(final + 2, 10)).Merge()
    ws.Range(ws.Cells(final + 2, 1), ws.Cells(final + 2, 10)).WrapText = True
    ws.Rows(final + 2).RowHeight = 34

    if not rows:
        return
    chart = ws.ChartObjects().Add(ws.Range("F6").Left, ws.Range("F6").Top, 550, 310)
    chart.Name = "DepartmentSummary"
    chart.Chart.ChartType = 51  # xlColumnClustered
    for column, label, rgb in (
        (2, "Αρχικά προγραμματισμένες", BLUE_RGB),
        (3, "Τελικά πραγματοποιημένες", RED_RGB),
    ):
        series = chart.Chart.SeriesCollection().NewSeries()
        series.Name = label
        series.XValues = ws.Range(ws.Cells(7, 1), ws.Cells(final - 1, 1))
        series.Values = ws.Range(ws.Cells(7, column), ws.Cells(final - 1, column))
        series.Format.Fill.ForeColor.RGB = rgb
    chart.Chart.HasTitle = True
    chart.Chart.ChartTitle.Text = MODE_LABELS[mode] + " | ΣΥΝΕΔΡΙΕΣ ΤΜΗΜΑΤΟΣ"
    chart.Chart.HasLegend = True
    chart.Chart.Axes(2).MinimumScale = 0
    chart.Chart.Axes(2).HasMajorGridlines = True
    ws.Activate()
    ws.Range("A1").Select()


def verify_management_native(file: Path, expected: dict) -> None:
    """Independently reopen the new XLSX read-only and reconcile all totals."""
    import win32com.client  # type: ignore[import-not-found]

    before_hash = hashlib.sha256(file.read_bytes()).hexdigest()
    excel = None
    book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        book = excel.Workbooks.Open(str(file.resolve()), UpdateLinks=0, ReadOnly=True)
        if not bool(book.ReadOnly) or bool(book.HasVBProject) or int(book.Worksheets.Count) != 4:
            raise ValueError("Management export read-only structure invalid")
        for mode in MODES:
            ws = book.Worksheets(MODE_LABELS[mode])
            if int(ws.Visible) != -1:
                raise ValueError(f"Management {mode} worksheet unexpectedly hidden")
            expected_rows = expected[mode]
            if int(ws.ChartObjects().Count) != (1 if expected_rows else 0):
                raise ValueError(f"Management {mode} chart count mismatch")
            planned_total = delivered_total = 0
            for offset, (period, planned, delivered, ratio) in enumerate(expected_rows):
                row = 7 + offset
                actual_period = str(ws.Cells(row, 1).Value2)
                actual_planned = int(ws.Cells(row, 2).Value2)
                actual_delivered = int(ws.Cells(row, 3).Value2)
                if (actual_period, actual_planned, actual_delivered) != (
                    period, planned, delivered,
                ):
                    raise ValueError(f"Management {mode} row {row} read-back mismatch")
                exported_ratio = ws.Cells(row, 4).Value2
                if ratio is None:
                    if str(exported_ratio) != "Μη διαθέσιμο":
                        raise ValueError("Zero denominator was not explicitly unavailable")
                elif not math.isclose(float(exported_ratio), ratio, abs_tol=1e-12):
                    raise ValueError(f"Management {mode} ratio mismatch")
                planned_total += planned
                delivered_total += delivered
            total_row = 7 + len(expected_rows)
            if (str(ws.Cells(total_row, 1).Value2) != "ΣΥΝΟΛΟ ΟΜΑΔΑΣ"
                    or int(ws.Cells(total_row, 2).Value2) != planned_total
                    or int(ws.Cells(total_row, 3).Value2) != delivered_total):
                raise ValueError(f"Management {mode} total read-back mismatch")
            total_ratio = ws.Cells(total_row, 4).Value2
            if planned_total == 0:
                if str(total_ratio) != "Μη διαθέσιμο":
                    raise ValueError("Team zero denominator marked incorrectly")
            elif not math.isclose(float(total_ratio),
                                  delivered_total / planned_total, abs_tol=1e-12):
                raise ValueError(f"Management {mode} total percentage mismatch")
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
    if hashlib.sha256(file.read_bytes()).hexdigest() != before_hash:
        raise ValueError("Read-only management validation altered the export")


def export_shareable(source: Path, output: Path) -> None:
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise ValueError("Existing .xlsm source required")
    if output.exists() or output.suffix.casefold() != ".xlsx":
        raise ValueError("Output must be a NEW .xlsx; never overwrite")
    if source.resolve() == output.resolve():
        raise ValueError("Source/destination collision")
    # Read report facts first. No spreadsheet data copied, only aggregation.
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    periods = department_periods(report_therapist_history(source))
    if not periods["overall"]:
        raise ValueError("No finalized history available to share")

    if sys.platform != "win32":
        raise RuntimeError("Native management XLSX export requires Windows Excel")
    import win32com.client  # type: ignore[import-not-found]

    excel = None
    book = None
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        book = excel.Workbooks.Add()
        while book.Worksheets.Count < 4:
            book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
        while book.Worksheets.Count > 4:
            book.Worksheets(book.Worksheets.Count).Delete()
        for idx, mode in enumerate(MODES, 1):
            _write_mode_sheet(book.Worksheets(idx), mode, periods[mode])
        book.Worksheets(1).Activate()
        # FileFormat 51 is macro-free XLSX. Fresh workbook with no source links.
        book.SaveAs(str(output.resolve()), FileFormat=51)
        book.Close(SaveChanges=False)
        book = None
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
    inspect_share_package(output)
    verify_management_native(output, periods)
    inspect_share_package(output)
    if hashlib.sha256(source.read_bytes()).hexdigest() != source_hash:
        raise ValueError("Source operational XLSM was modified during export")


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a shareable department-only XLSX")
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    try:
        export_shareable(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        print("Never share an unverified export. Do not overwrite or retry the same destination.")
        return 2
    print("MANAGEMENT-ONLY EXPORT VERIFIED OK")
    print("Four aggregate modes, four sheets; no VBA, no external links or copied therapist records")
    print("Source workbook unchanged by export.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
