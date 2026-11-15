"""Read-only structural check; no COM / workbook writes / patient row output."""
from __future__ import annotations

import argparse
from pathlib import Path

from rehab_excel.afternoon_native_preflight import (
    AfternoonNativePreflightError,
    inspect_afternoon_native_candidate,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Inspect only XLSM sheet roster/VBA, not patient data or approval"
    )
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = inspect_afternoon_native_candidate(args.source)
    except AfternoonNativePreflightError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    print("AFTERNOON XLSM STRUCTURAL PREFLIGHT OK")
    print(f"Sheet count: {result.sheet_count}")
    print("VBA binary present: YES")
    print("Afternoon sheet already present: NO")
    print("NO PATIENT DATA READ OR WRITTEN")
    print("Not a synthetic-provenance or native installation approval")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
