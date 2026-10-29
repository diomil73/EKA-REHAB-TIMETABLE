from __future__ import annotations

import argparse
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Quarantined experimental Manual Replacement Operator V2 repair. "
            "Do not mutate disposable smoke workbooks through live-VBA marker patching."
        )
    )
    parser.add_argument("workbook", type=Path)
    parser.parse_args()

    print(
        "SAFETY STOP: Manual Replacement Operator V2 live-VBA repair is quarantined. "
        "Rebuild a fresh disposable smoke from canonical source and implement UI changes "
        "through the normal generator/build pipeline instead."
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
