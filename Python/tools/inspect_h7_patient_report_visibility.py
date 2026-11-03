from __future__ import annotations

import argparse
from pathlib import Path
import sys


def _is_numeric_id(value: object, number: int) -> bool:
    label = str(value if value is not None else "").strip()
    return label.isascii() and label.isdecimal() and int(label) == number


def inspect_patient_excel_visibility(path: Path) -> int:
    """Examine saved patient-report rows without saving or changing the workbook."""
    if sys.platform != "win32" or not path.is_file():
        print("SAFETY STOP: Windows Excel and existing workbook required")
        return 2
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 required")
        return 2
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        ws = wb.Worksheets("PATIENT_HISTORY_REPORT")
        last_row = int(ws.Cells(ws.Rows.Count, 3).End(-4162).Row)
        print("H7 PATIENT REPORT VISIBILITY READ-ONLY AUDIT")
        print("Workbook:", wb.FullName)
        print("ReadOnly:", bool(wb.ReadOnly))
        print("Saved mode label:", str(ws.Range("A2").Value2))
        print("AutoFilterMode:", bool(ws.AutoFilterMode))
        print("FilterMode:", bool(ws.FilterMode))
        print("ScrollArea:", str(ws.ScrollArea))
        print("Last populated patient row:", last_row)
        for patient_no in (15, 18, 19, 20, 21, 22):
            locations = []
            for row in range(6, last_row + 1):
                if _is_numeric_id(ws.Cells(row, 3).Value2, patient_no):
                    locations.append(
                        f"row={row} hidden={bool(ws.Rows(row).Hidden)} "
                        f"height={float(ws.Rows(row).RowHeight):.1f} "
                        f"text={str(ws.Cells(row, 3).Text)!r}"
                    )
            print(f"ID {patient_no:02d}:", " | ".join(locations) if locations else "NOT IN SAVED REPORT SHEET")
        print("H7 PATIENT REPORT VISIBILITY AUDIT COMPLETE; workbook unchanged")
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit saved patient report view nonmutating")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return inspect_patient_excel_visibility(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
