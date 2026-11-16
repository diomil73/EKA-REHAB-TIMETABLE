"""Private outpatient export audit, using current Diomil source schemas only.

The tool deliberately never imports, creates a demo, prints names/IDs or
guesses an outpatient source from the inpatient A:V programme.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

PYTHON = Path(__file__).resolve().parents[1]
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.trial_private_outpatient_audit import (  # noqa: E402
    inspect_private_outpatient_source,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Read-only actual outpatient PATIENTS + OUTPATIENT_SCHEDULE audit"
    )
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = inspect_private_outpatient_source(args.source)
    except Exception as error:
        # Source exceptions may contain private patient names and paths.
        print(f"SAFETY STOP: {type(error).__name__}; no outpatient records imported.")
        return 2
    print("DIOMIL PRIVATE OUTPATIENT SOURCE AUDIT OK")
    print(f"Explicitly typed outpatient patients: {result.outpatient_patients}")
    print(f"Existing-schema recurring treatments: {result.recurring_treatments}")
    print(f"Confirmed Monday-Friday placements: {result.weekday_placements}")
    print(f"Distinct existing treatment categories: {result.distinct_treatments}")
    print("Source unchanged: YES")
    print("Native DEMO outpatient import: PENDING; this audit grants no admission.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
