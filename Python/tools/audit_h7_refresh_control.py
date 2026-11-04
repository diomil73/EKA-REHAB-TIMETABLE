from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys


def _safe_shape_value(shape, attribute: str) -> str:
    try:
        value = shape
        for part in attribute.split("."):
            value = getattr(value, part)
        return "" if value is None else str(value)
    except Exception:
        return ""


def audit(path: Path) -> int:
    """Read-only inspection of RefreshAll assignment, VBA source and refresh dependencies."""
    if sys.platform != "win32" or not path.is_file() or path.suffix.lower() != ".xlsm":
        print("SAFETY STOP: existing XLSM on Windows is required")
        return 2
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 is required")
        return 2

    excel = None
    workbook = None
    errors = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        ws = workbook.Worksheets("THERAPIST_DAILY")

        print("H7 REFRESH CONTROL READ-ONLY AUDIT")
        print("Workbook:", workbook.FullName)
        print("ReadOnly:", bool(workbook.ReadOnly))
        matches = []
        for index in range(1, int(ws.Shapes.Count) + 1):
            shape = ws.Shapes(index)
            action = _safe_shape_value(shape, "OnAction")
            caption = _safe_shape_value(shape, "TextFrame2.TextRange.Text")
            if "RefreshAll" in action or "ΑΝΑΝΕΩΣΗ" in caption:
                print(
                    "REFRESH SHAPE:",
                    f"name={shape.Name!s}",
                    f"caption={caption!r}",
                    f"type={int(shape.Type)}",
                    f"macro={action!r}",
                )

        try:
            components = workbook.VBProject.VBComponents
            print("VBA component count:", int(components.Count))
            for index in range(1, int(components.Count) + 1):
                component = components.Item(index)
                module = component.CodeModule
                count = int(module.CountOfLines)
                if not count:
                    continue
                source = str(module.Lines(1, count))
                lines = source.splitlines()
                for line_no, line in enumerate(lines):
                    if re.search(r"\bRefreshAll\b", line, flags=re.IGNORECASE):
                        matches.append((str(component.Name), line_no + 1, lines[max(0, line_no - 3):line_no + 5]))
            print("RefreshAll source references:", len(matches))
            for component, line_no, excerpt in matches:
                print(f"VBA MATCH {component}:{line_no}")
                for line in excerpt:
                    print("  " + line)
        except Exception as exc:
            errors.append(f"VBA project source unavailable: {type(exc).__name__}: {exc}")
            print("VBA source inspection unavailable:", errors[-1])

        try:
            connections = int(workbook.Connections.Count)
        except Exception:
            connections = -1
        try:
            pivot_caches = int(workbook.PivotCaches().Count)
        except Exception:
            pivot_caches = -1
        query_tables = 0
        pivot_tables = 0
        for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
            sheet = workbook.Worksheets(sheet_index)
            try:
                query_tables += int(sheet.QueryTables.Count)
            except Exception:
                pass
            try:
                pivot_tables += int(sheet.PivotTables().Count)
            except Exception:
                pass
        print("Workbook connections:", connections)
        print("Pivot caches:", pivot_caches)
        print("Worksheet query tables:", query_tables)
        print("Pivot tables:", pivot_tables)
        print("VBA RefreshAll procedure found:", any(
            re.search(r"\b(Sub|Function)\s+RefreshAll\b", "\n".join(lines), re.IGNORECASE)
            for _, _, lines in matches
        ))
        print("Workbook unchanged: True")
        if errors:
            print("H7 REFRESH AUDIT INCOMPLETE")
            return 3
        print("H7 REFRESH AUDIT COMPLETE")
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only inspect RefreshAll macro and workbook refresh dependencies")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return audit(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
