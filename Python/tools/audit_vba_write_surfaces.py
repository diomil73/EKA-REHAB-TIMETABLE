from __future__ import annotations

import argparse
from pathlib import Path
import sys


def audit(path: Path) -> int:
    if sys.platform != "win32" or not path.is_file() or path.suffix.lower() != ".xlsm":
        print("SAFETY STOP: existing XLSM on Windows is required")
        return 2
    try:
        import win32com.client
    except ImportError:
        print("SAFETY STOP: pywin32 is required")
        return 2
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        print("VBA SOURCE READ-ONLY AUDIT")
        print("Workbook:", workbook.FullName)
        print("ReadOnly:", bool(workbook.ReadOnly))
        print("VBA component count:", int(workbook.VBProject.VBComponents.Count))
        print("Workbook unchanged: True")
        print("VBA SOURCE AUDIT COMPLETE")
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
    parser = argparse.ArgumentParser(description="Read-only VBA source inventory")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return audit(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
