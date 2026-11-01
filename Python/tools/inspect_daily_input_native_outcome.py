from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook


def _font_color(cell) -> str:
    color = cell.font.color
    if color is None:
        return "<none>"
    if color.type == "rgb":
        return str(color.rgb)
    return f"{color.type}:{color.indexed if color.type == 'indexed' else color.theme}"


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only inspection of DAILY_INPUT native post-run state")
    parser.add_argument("--workbook", required=True, type=Path)
    args = parser.parse_args()

    path = args.workbook.resolve()
    if not path.exists():
        print(f"SAFETY STOP: workbook not found: {path}")
        return 2

    wb = load_workbook(path, read_only=False, data_only=False, keep_vba=True)
    try:
        print("WORKBOOK:", path)

        if "DAILY_INPUT" in wb.sheetnames:
            ws = wb["DAILY_INPUT"]
            print("DAILY_INPUT_B2:", ws["B2"].value)
            nonempty = []
            for row in range(1, min(ws.max_row, 40) + 1):
                vals = [ws.cell(row, col).value for col in range(1, min(ws.max_column, 12) + 1)]
                if any(v not in (None, "") for v in vals):
                    nonempty.append((row, vals))
            print("DAILY_INPUT_NONEMPTY_ROWS:")
            for row, vals in nonempty:
                print(f"  row {row}: {vals}")

        if "THERAPIST_DAILY" in wb.sheetnames:
            ws = wb["THERAPIST_DAILY"]
            for addr in ("B7", "D2", "H2", "I13", "J13", "E14", "E15", "G6", "J16", "F17"):
                cell = ws[addr]
                print(
                    f"THERAPIST_DAILY_{addr}: value={cell.value!r} strike={cell.font.strike!r} "
                    f"italic={cell.font.italic!r} color={_font_color(cell)} fill={cell.fill.fgColor.rgb!r}"
                )

        if "MASTER_SCHEDULE" in wb.sheetnames:
            ws = wb["MASTER_SCHEDULE"]
            for addr in ("D8", "M8", "D9", "M9"):
                cell = ws[addr]
                print(
                    f"MASTER_{addr}: value={cell.value!r} strike={cell.font.strike!r} "
                    f"italic={cell.font.italic!r} color={_font_color(cell)} fill={cell.fill.fgColor.rgb!r}"
                )

        if "DAILY_CHANGE_LEDGER" in wb.sheetnames:
            ws = wb["DAILY_CHANGE_LEDGER"]
            print("LEDGER_ROWS:")
            for row in range(1, ws.max_row + 1):
                vals = [ws.cell(row, col).value for col in range(1, min(ws.max_column, 12) + 1)]
                if any(v not in (None, "") for v in vals):
                    print(f"  row {row}: {vals}")
        else:
            print("LEDGER_ROWS: <sheet missing>")

        return 0
    finally:
        wb.close()


if __name__ == "__main__":
    raise SystemExit(main())
