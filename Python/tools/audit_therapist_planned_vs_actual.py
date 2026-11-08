from __future__ import annotations

"""Read-only explanation of days where delivered sessions exceed initial plan.

Only current-final day snapshots are read through the canonical report tool.
No XLSM modifications, no retries or auto repairs. Positive deltas are
flagged for review rather than silently clipped or reclassified.
"""

import argparse
from collections import defaultdict
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.report_therapist_history import report_therapist_history


def positive_deltas(reports, therapist: str | None = None) -> list[dict[str, object]]:
    normalized = therapist.strip().casefold() if therapist else None
    flagged = []
    observed = set()
    for report in reports:
        key = (report.service_date, report.therapist_id.casefold())
        if key in observed:
            raise ValueError("Duplicate current-final therapist/date in diagnostic")
        observed.add(key)
        if normalized and report.therapist_id.casefold() != normalized:
            continue
        planned = int(report.planned_sessions)
        delivered = int(report.final_active_sessions)
        if planned < 0 or delivered < 0:
            raise ValueError("Negative session count in finalized history")
        if delivered > planned:
            flagged.append({
                "date": report.service_date.isoformat(),
                "therapist": report.therapist_id,
                "initial_planned": planned,
                "final_delivered": delivered,
                "delta": delivered - planned,
                "replacements": int(report.replacements_received),
                "approved_extra": int(report.approved_shared_or_over_capacity_events),
                "released": int(report.released_by_cancellation_or_no_show),
                "suspensions": int(report.treatment_suspensions_affecting_load),
            })
    return sorted(flagged, key=lambda row: (row["date"], str(row["therapist"]).casefold()))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read-only investigation of planned vs delivered therapist sessions."
    )
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--therapist", default=None, help="Optional exact therapist label")
    args = parser.parse_args()
    if not args.workbook.is_file() or args.workbook.suffix.casefold() != ".xlsm":
        print("SAFETY STOP: existing .xlsm required; no workbook changed")
        return 2
    try:
        reports = report_therapist_history(args.workbook)
        flagged = positive_deltas(reports, args.therapist)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2

    print("PLANNED VS DELIVERED READ-ONLY AUDIT OK")
    print(f"Finalized therapist/day records: {len(reports)}")
    print(f"Positive daily deltas: {len(flagged)}")
    for row in flagged:
        print(
            f"{row['date']} | {row['therapist']} | "
            f"initial={row['initial_planned']} | delivered={row['final_delivered']} "
            f"| delta=+{row['delta']} | replacements={row['replacements']} "
            f"| approved_extra={row['approved_extra']} "
            f"| released={row['released']} | suspensions={row['suspensions']}"
        )
    print("NOTE: Positive delta is NOT proof of a particular cause.")
    print("Workbook unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
