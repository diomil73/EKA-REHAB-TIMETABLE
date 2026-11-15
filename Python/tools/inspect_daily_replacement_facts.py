from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_replacement_facts import (  # noqa: E402
    DailyReplacementFactError,
    read_daily_replacement_fact_rows,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="List persisted daily replacement facts.")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        facts = read_daily_replacement_fact_rows(
            args.workbook,
            include_restored=True,
        )
    except DailyReplacementFactError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DAILY REPLACEMENT FACT INVENTORY")
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"Facts: {len(facts)}")
    for index, fact in enumerate(facts, start=1):
        print(
            f"{index}. id={fact.replacement_id!r} | status={fact.status} | "
            f"patient={fact.patient_id!r} | {fact.original_therapist_id} -> "
            f"{fact.replacement_therapist_id} {fact.replacement_time.strftime('%H:%M')} | "
            f"date={fact.replacement_date.isoformat()}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
