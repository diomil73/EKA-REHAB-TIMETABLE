"""Run a private, read-only source/review provenance check.

Prints aggregate counts and safe issue codes only. Never print source paths,
patient names, PatientIDs, ward/doctor values, or raw cell contents.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.trial_source_alignment import (  # noqa: E402
    TrialSourceAlignmentError,
    audit_trial_source_alignment,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="READ-ONLY private A:V programme/review alignment audit"
    )
    parser.add_argument("--programme", required=True, type=Path)
    parser.add_argument("--review", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = audit_trial_source_alignment(
            programme=args.programme, review=args.review
        )
    except (TrialSourceAlignmentError, OSError, ValueError) as exc:
        print(f"SAFETY STOP: source alignment could not complete ({type(exc).__name__}).")
        return 2

    print("DIOMIL PRIVATE SOURCE ALIGNMENT AUDIT")
    print(f"Patient rows checked: {result.patient_rows_checked}")
    print(f"Weekly rows checked: {result.weekly_rows_checked}")
    print(f"Source references checked: {result.source_references_checked}")
    print(f"Rows with missing ward/doctor context: {result.missing_patient_context_rows}")
    print(f"Patient identity mismatches: {result.identity_mismatches}")
    print(f"Weekly identity mismatches: {result.schedule_identity_mismatches}")
    print(f"Source reference mismatches: {result.source_reference_mismatches}")
    print(f"Unsupported specialties: {result.unsupported_specialties}")
    print(f"Inputs unchanged: {'YES' if result.inputs_unchanged else 'NO'}")
    for blocker in result.blockers:
        print(f"REVIEW REQUIRED: {blocker}")
    if not result.aligned:
        print("SAFETY STOP: source alignment is not complete; no import is authorized.")
        return 2
    print("SOURCE ALIGNMENT GREEN ONLY; IDENTITY APPROVAL AND IMPORT NOT AUTHORIZED.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
