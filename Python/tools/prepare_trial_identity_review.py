"""Create a local-only review copy of patient ID suggestions.

Use ONLY on institution-approved private storage. Prints counts, never patient
names/IDs. No changes to source PROGRAMMA, old demo or active MASTER.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "Python"))

from rehab_excel.trial_identity_review import (  # noqa: E402
    TrialIdentityReviewError, create_local_identity_review,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Diomil local patient identity review")
    parser.add_argument("--staging", required=True, type=Path)
    parser.add_argument("--cc", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        counts = create_local_identity_review(
            staging_path=args.staging, old_demo_path=args.cc, output_path=args.output,
        )
    except (TrialIdentityReviewError, OSError, ValueError, KeyError) as exc:
        # Hide raw paths and patient identity material from console/logs.
        print(f"SAFETY STOP: Patient identity review failed ({type(exc).__name__}).")
        return 2
    print("READ-ONLY INPUTS: patient review copy created locally.")
    print(f"Source patient rows: {counts['source_rows']}")
    print(f"Unique old-ID suggestions (UNCONFIRMED): {counts['exact_unconfirmed_suggestions']}")
    print(f"Rows requiring individual identity review: {counts['requires_individual_review']}")
    print("No PatientIDs were auto-confirmed and no .xlsm was modified.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
