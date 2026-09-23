from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]

TARGET_HEADERS = (
    "ΦΘ_Ώρα",
    "ΦΘ_Ημέρες",
    "ΦΘ_Θεραπευτής",
    "ΕΦΑ_Ώρα",
    "ΕΦΑ_Ημέρες",
)


def _header_map(ws) -> dict[str, int]:
    last_col = int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)  # xlToLeft
    result: dict[str, int] = {}
    for col in range(1, last_col + 1):
        value = ws.Cells(1, col).Value
        if value is None:
            continue
        text = str(value).strip()
        if text:
            result[text] = col
    return result


def _validation_snapshot(cell) -> str:
    try:
        validation = cell.Validation
        parts = [
            f"Type={validation.Type}",
            f"Formula1={validation.Formula1!r}",
            f"Formula2={validation.Formula2!r}",
            f"InCellDropdown={validation.InCellDropdown}",
        ]
        return ", ".join(parts)
    except Exception as exc:
        return f"NO/UNREADABLE VALIDATION: {exc}"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read existing PATIENT_PLANNER Excel validations without modifying the workbook."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    args = parser.parse_args()
    source = args.source.resolve()

    if sys.platform != "win32":
        print("SAFETY STOP: requires Windows with Microsoft Excel.")
        return 2
    if not source.exists():
        print(f"SAFETY STOP: workbook not found: {source}")
        return 2

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 is required.")
        return 2

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True)
        ws = workbook.Worksheets("PATIENT_PLANNER")
        headers = _header_map(ws)

        print("PATIENT_PLANNER EXISTING VALIDATIONS")
        for header in TARGET_HEADERS:
            col = headers.get(header)
            if col is None:
                print(f"{header}: HEADER NOT FOUND")
                continue
            print(f"{header}: col {col}")
            for row in (2, 3, 10):
                cell = ws.Cells(row, col)
                print(f"  row {row}: {_validation_snapshot(cell)}")

        print("WORKBOOK NAMES")
        count = int(workbook.Names.Count)
        for index in range(1, count + 1):
            try:
                item = workbook.Names.Item(index)
                print(f"  {item.Name} -> {item.RefersTo}")
            except Exception as exc:
                print(f"  [name {index} unreadable: {exc}]")
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
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


if __name__ == "__main__":
    raise SystemExit(main())
