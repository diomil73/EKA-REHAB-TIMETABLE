from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_change_ledger_live_protection_patch import (  # noqa: E402
    install_daily_change_ledger_live_protection,
)
from rehab_excel.manual_operator_v2_source_normalizer import (  # noqa: E402
    normalize_manual_operator_v2_sources,
)
from rehab_excel.manual_replacement_operator_v2_patch import (  # noqa: E402
    install_manual_operator_v2,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Install manual replacement operator V2 into a disposable .xlsm workbook.")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        install_daily_change_ledger_live_protection(args.workbook)
        normalize_manual_operator_v2_sources(args.workbook)
        install_manual_operator_v2(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("MANUAL REPLACEMENT OPERATOR V2 REPAIR OK")
    print(f"Workbook: {args.workbook.resolve()}")
    print("Protected ledger rendering: installed")
    print("Live VBA source normalization: installed")
    print("Treatment suspension action: installed")
    print("Explicit provider/time selection: installed")
    print("Previous/next operator navigation: installed")
    print("Change-current decision: installed")
    print("Automatic remainder route: installed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
