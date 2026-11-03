from __future__ import annotations

import argparse
from pathlib import Path
import sys


def inspect(path: Path) -> int:
    if sys.platform != "win32":
        print("SAFETY STOP: Windows Excel required")
        return 2
    if not path.is_file():
        print(f"SAFETY STOP: Workbook missing: {path}")
        return 2
    import win32com.client  # type: ignore[import-not-found]

    excel = win32com.client.DispatchEx("Excel.Application")
    wb = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        print("UNIFIED DAY SAVE READ-ONLY AUDIT")
        print(f"Workbook: {wb.FullName}")
        for sheet in ("MASTER_SCHEDULE", "DAILY_INPUT"):
            ws = wb.Worksheets(sheet)
            print(f"SHEET {sheet}:")
            matches = []
            for i in range(1, int(ws.Shapes.Count) + 1):
                s = ws.Shapes(i)
                name = str(s.Name)
                if "SaveOperationalDay" in name or "CloseDay" in name or "DayAmendment" in name or name == "ekaToolbar_ExitApp":
                    matches.append(name)
                    print(f"  {name}: OnAction={s.OnAction!s}; Left={float(s.Left):.1f}; Top={float(s.Top):.1f}; Width={float(s.Width):.1f}; Visible={bool(s.Visible)}")
            if not matches:
                print("  NO MATCHING SHAPES")
        try:
            cm = wb.VBProject.VBComponents("modUnifiedDaySave").CodeModule
            code = cm.Lines(1, cm.CountOfLines)
            print(f"DISPATCHER: {'PRESENT' if 'Public Sub SaveOperationalDay()' in code else 'INVALID'}")
        except Exception as exc:
            print(f"DISPATCHER: MISSING ({type(exc).__name__})")
        print("READ-ONLY AUDIT COMPLETE; no workbook changes")
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: read-only audit failed: {exc}")
        return 2
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only inspection of persisted unified day-save controls")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return inspect(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
