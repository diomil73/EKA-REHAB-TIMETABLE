from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_change_ledger_live_protection_patch import (  # noqa: E402
    DailyChangeLedgerLiveProtectionPatchError,
    install_daily_change_ledger_live_protection,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Repair THERAPIST_DAILY ledger rendering protection in an existing .xlsm workbook."
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        install_daily_change_ledger_live_protection(args.workbook)
    except DailyChangeLedgerLiveProtectionPatchError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DAILY CHANGE LEDGER LIVE PROTECTION REPAIR OK")
    print(f"Workbook: {args.workbook.resolve()}")
    print("THERAPIST_DAILY render bracket: installed")
    print("Protection restored on success/error: verified by patch contract")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
