from __future__ import annotations

import argparse
from pathlib import Path
import sys


MAX_CELLS_PER_SHEET = 150_000
MAX_UNLOCKED_RANGES_PRINTED = 120


def column_letter(number: int) -> str:
    if number < 1:
        raise ValueError("Excel columns are 1-based")
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def selection_label(value: int) -> str:
    return {
        0: "xlNoRestrictions",
        1: "xlUnlockedCells",
        2: "xlNoSelection",
    }.get(value, f"unknown({value})")


def compress_unlocked_columns(row_number: int, columns: list[int]) -> list[str]:
    """Compress sorted column numbers into horizontal A1 ranges for one row."""
    if not columns:
        return []
    ordered = sorted(set(columns))
    ranges: list[str] = []
    start = previous = ordered[0]
    for column in ordered[1:]:
        if column != previous + 1:
            left = column_letter(start)
            right = column_letter(previous)
            ranges.append(f"{left}{row_number}" if start == previous else f"{left}{row_number}:{right}{row_number}")
            start = column
        previous = column
    left = column_letter(start)
    right = column_letter(previous)
    ranges.append(f"{left}{row_number}" if start == previous else f"{left}{row_number}:{right}{row_number}")
    return ranges


def _safe_bool(obj, attribute: str):
    try:
        return bool(getattr(obj, attribute))
    except Exception:
        return None


def _safe_value(obj, attribute: str):
    try:
        return getattr(obj, attribute)
    except Exception:
        return None


def _audit_protected_used_range(ws, used, rows: int, columns: int, count: int) -> tuple[int, list[str], bool]:
    unlocked_count = 0
    unlocked_ranges: list[str] = []
    complete = True
    if count > MAX_CELLS_PER_SHEET:
        return unlocked_count, [f"SCAN_SKIPPED: {count} cells exceeds limit {MAX_CELLS_PER_SHEET}"], False

    first_row = int(used.Row)
    first_column = int(used.Column)
    for row_offset in range(1, rows + 1):
        unlocked_columns: list[int] = []
        for column_offset in range(1, columns + 1):
            try:
                cell = used.Cells.Item(row_offset, column_offset)
                if not bool(cell.Locked):
                    unlocked_count += 1
                    unlocked_columns.append(first_column + column_offset - 1)
            except Exception as exc:
                complete = False
                unlocked_ranges.append(
                    f"SCAN_ERROR at {column_letter(first_column + column_offset - 1)}"
                    f"{first_row + row_offset - 1}: {type(exc).__name__}"
                )
                if len(unlocked_ranges) >= MAX_UNLOCKED_RANGES_PRINTED:
                    return unlocked_count, unlocked_ranges, False
        unlocked_ranges.extend(
            compress_unlocked_columns(first_row + row_offset - 1, unlocked_columns)
        )
        if len(unlocked_ranges) > MAX_UNLOCKED_RANGES_PRINTED:
            unlocked_ranges = unlocked_ranges[:MAX_UNLOCKED_RANGES_PRINTED]
            unlocked_ranges.append("... output truncated")
            break
    return unlocked_count, unlocked_ranges, complete


def audit(path: Path) -> int:
    """Read-only Excel protection inventory. Never calls Protect, Unprotect or Save."""
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
    incomplete: list[str] = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)

        print("WORKBOOK PROTECTION READ-ONLY AUDIT")
        print("Workbook:", workbook.FullName)
        print("ReadOnly:", bool(workbook.ReadOnly))
        print("ProtectStructure:", _safe_bool(workbook, "ProtectStructure"))
        print("ProtectWindows:", _safe_bool(workbook, "ProtectWindows"))
        print("Worksheet count:", int(workbook.Worksheets.Count))

        for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
            ws = workbook.Worksheets(sheet_index)
            name = str(ws.Name)
            contents = _safe_bool(ws, "ProtectContents")
            drawing = _safe_bool(ws, "ProtectDrawingObjects")
            scenarios = _safe_bool(ws, "ProtectScenarios")
            visibility = _safe_value(ws, "Visible")
            enable_selection = _safe_value(ws, "EnableSelection")
            used = ws.UsedRange
            rows = int(used.Rows.Count)
            columns = int(used.Columns.Count)
            count = int(used.Cells.CountLarge)
            locked_state = _safe_value(used, "Locked")
            locked_label = "MIXED/NULL" if locked_state is None else str(bool(locked_state))

            print(
                f"SHEET {name}:",
                f"Visible={visibility}",
                f"ProtectContents={contents}",
                f"ProtectDrawingObjects={drawing}",
                f"ProtectScenarios={scenarios}",
                f"EnableSelection={selection_label(int(enable_selection)) if enable_selection is not None else 'UNKNOWN'}",
                f"UsedRange={str(used.Address)}",
                f"UsedCells={count}",
            )
            print("UsedRange.Locked", locked_label)

            protection = _safe_value(ws, "Protection")
            if protection is not None:
                for option in (
                    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
                    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
                    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
                    "AllowFiltering", "AllowUsingPivotTables",
                ):
                    value = _safe_value(protection, option)
                    if value is not None:
                        print(f"  {option}={bool(value)}")

            if contents is not True:
                print(f"  SAFETY FINDING: worksheet cell contents are not protected ({name})")
                continue

            if locked_state is False:
                print("  UNLOCKED CELLS: entire UsedRange is unlocked")
                continue
            if locked_state is True:
                print("  UNLOCKED CELLS: none in UsedRange")
                continue

            unlocked_count, ranges, complete = _audit_protected_used_range(
                ws, used, rows, columns, count
            )
            print(f"  Unlocked cells in UsedRange: {unlocked_count}")
            for address in ranges:
                print("  UNLOCKED:", address)
            if not complete:
                incomplete.append(name)
                print(f"  SCAN INCOMPLETE: {name}")

        print("Workbook unchanged: True")
        if incomplete:
            print("WORKBOOK PROTECTION AUDIT INCOMPLETE:", ", ".join(incomplete))
            return 3
        print("WORKBOOK PROTECTION AUDIT COMPLETE")
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
    parser = argparse.ArgumentParser(
        description="Read-only audit of worksheet protection, locked ranges and workbook structure"
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return audit(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
