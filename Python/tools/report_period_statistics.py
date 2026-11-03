from __future__ import annotations

import argparse
from datetime import date
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.period_statistics import (  # noqa: E402
    period_statistics_dicts,
    read_period_statistics,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only finalized-day statistics report")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--period", required=True, choices=("day", "week", "month"))
    parser.add_argument("--start", type=date.fromisoformat)
    parser.add_argument("--end", type=date.fromisoformat)
    args = parser.parse_args()
    if not args.workbook.is_file():
        print("SAFETY STOP: Workbook does not exist")
        return 2
    if args.start and args.end and args.start > args.end:
        print("SAFETY STOP: start cannot be after end")
        return 2
    try:
        rows = read_period_statistics(
            args.workbook, period=args.period, start=args.start, end=args.end
        )
        print(json.dumps({
            "ok": True,
            "period": args.period,
            "report_rows": period_statistics_dicts(rows),
            "workbook_unchanged": True,
        }, ensure_ascii=False, indent=2))
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
