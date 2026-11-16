"""Privacy-safe command: inspect only F/U/V of an actual private A:V workbook.

This is NOT a converter or special-treatment approval. Never print patient
identifiers, source paths, clinic, doctor or raw cell content.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.trial_special_fields_audit import (  # noqa: E402
    inspect_real_manual_special_fields,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="READ-ONLY private A:V reclining/afternoon/pool blocker inventory"
    )
    parser.add_argument("--programme", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = inspect_real_manual_special_fields(args.programme)
    except Exception as exc:
        # Raw exceptions may contain clinical values or private file paths.
        print(f"SAFETY STOP: {type(exc).__name__}; source not admitted.")
        return 2
    print("DIOMIL ACTUAL MANUAL A:V SPECIAL-FIELD INVENTORY")
    print(f"Source patient rows: {result.patient_rows}")
    print(f"F reclining, 30-min duration: {result.reclining_rows}")
    print(f"F clocks readable, weekdays STILL REQUIRE REVIEW: {result.reclining_clock_parseable}")
    print(f"F clocks ambiguous / formulas: {result.reclining_clock_review}")
    print(f"U afternoon unimported records: {result.afternoon_rows}")
    print(f"V morning pool records: {result.pool_rows}")
    print(f"V clocks readable, weekdays STILL REQUIRE REVIEW: {result.pool_clock_parseable}")
    print(f"V clocks ambiguous / formulas: {result.pool_clock_review}")
    print(f"Special-column formula cells: {result.formula_fields}")
    print(f"Special rows missing patient name: {result.orphan_rows}")
    print(f"Total special field blockers: {result.unresolved_fields}")
    print("Original workbook unchanged: YES")
    if result.unresolved_fields:
        print("REVIEW REQUIRED: no special treatment admitted or silently omitted.")
        return 3
    print("SPECIAL F/U/V CELLS EMPTY: full source admission still separately required.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
