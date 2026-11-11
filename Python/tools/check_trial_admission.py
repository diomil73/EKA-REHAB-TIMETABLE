"""Read-only gate for the current Diomil exploratory-program staging workbook.

Example:
 python Python/tools/check_trial_admission.py --review "private-review.xlsx"

Prints only aggregate counts, never names or patient IDs; never writes an xlsm.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

from rehab_excel.trial_admission_gate import TrialAdmissionError, evaluate_trial_admission


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only Diomil pilot source safety gate")
    parser.add_argument("--review", required=True, type=Path)
    args = parser.parse_args(argv)
    if not args.review.is_file():
        print("SAFETY STOP: Review workbook does not exist.")
        return 2
    try:
        status = evaluate_trial_admission(args.review)
    except (TrialAdmissionError, ValueError, OSError, KeyError) as exc:
        # Never print raw cell values or file paths in a user-accessible log.
        print(f"SAFETY STOP: Source admission validation failed ({type(exc).__name__}).")
        return 2

    print(f"Candidate patient rows: {status.candidate_patients}")
    print(f"Confirmed patient IDs: {status.verified_patient_ids}")
    print(f"Proposed weekday slots: {status.weekly_proposals}")
    print(f"Unadmitted weekday slots: {status.unadmitted_weekly_rows}")
    print(f"Unresolved source rows: {status.unresolved_source_rows}")
    print(f"Overlapping source rows: {status.overlap_rows}")
    print(f"Duplicate confirmed IDs: {status.duplicate_ids}")
    if not status.ready:
        print("SAFETY STOP: No workbook import authorized.")
        for blocker in status.blockers:
            print(" - " + blocker)
        return 2
    print("PREVIEW PASS: Staging may proceed to independent COM/native Excel checks.")
    print("No workbook was modified. This is NOT an Excel import or release.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
