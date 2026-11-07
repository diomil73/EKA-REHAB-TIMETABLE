from __future__ import annotations

"""READ-ONLY diagnostics for frozen failed synthetic demo, never repair."""

import argparse
import hashlib
from pathlib import Path
import sys

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))


SHEETS = (
    "FINALIZED_DAY_SNAPSHOTS",
    "THERAPIST_DAY_HISTORY",
    "PATIENT_SESSION_HISTORY",
    "DAY_HISTORY_EVENTS",
    "THERAPIST_ATTENDANCE_HISTORY",
)


def audit(path: Path) -> None:
    if "_DEMO_" not in path.name.upper():
        raise ValueError("Diagnostic only accepts _DEMO_ workbook")
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        for name in SHEETS:
            if name not in wb.sheetnames:
                print(f"{name}: MISSING")
                continue
            ws = wb[name]
            last = ws.max_row
            filled = 0
            partial = []
            for index, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if all(value is None or str(value).strip() == "" for value in values):
                    continue
                filled += 1
                if values[0] is None or not str(values[0]).strip():
                    partial.append((index, tuple(str(value)[:55] for value in values[:6])))
            print(f"{name}: max_row={last} nonempty_rows={filled} blank_id_partial_rows={len(partial)}")
            for row, short in partial[:6]:
                print(f"  BAD ROW {row}: {short!r}")
            if name == "FINALIZED_DAY_SNAPSHOTS":
                for index in range(max(2, last - 5), last + 1):
                    values = tuple(ws.cell(index, col).value for col in range(1, 7))
                    print(f"  ROW {index}: {values!r}")
    finally:
        wb.close()
    after = hashlib.sha256(path.read_bytes()).hexdigest()
    print(f"Workbook unchanged: {before == after}")
    if before != after:
        raise RuntimeError("Read-only diagnostic changed workbook")


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only failed H8 demo seed diagnosis")
    parser.add_argument("--workbook", required=True, type=Path)
    args = parser.parse_args()
    try:
        audit(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
