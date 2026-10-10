from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.workbook_scheduling_candidates import (  # noqa: E402
    workbook_therapist_candidates,
)


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def render_candidate_report(
    workbook: Path,
    *,
    patient_id: str,
    treatment: str,
    target_date: date,
) -> str:
    candidates = workbook_therapist_candidates(
        workbook,
        patient_id=patient_id,
        treatment=treatment,
        target_date=target_date,
    )

    lines = [
        "SCHEDULING CANDIDATES",
        f"Workbook: {workbook}",
        f"PatientID: {patient_id}",
        f"Treatment: {treatment}",
        f"Date: {target_date.isoformat()}",
        "",
    ]

    if not candidates:
        lines.append("No feasible therapist/time candidates found.")
        return "\n".join(lines)

    for index, candidate in enumerate(candidates, start=1):
        capacity = f"{candidate.active_timeslots}/{candidate.max_daily_timeslots} slots"
        lines.append(
            f"{index}. {candidate.display_name} | load {candidate.active_sessions} | {capacity}"
        )
        for slot in candidate.available_slots:
            lines.append(f"   - {slot.strftime('%H:%M')}")
        lines.append("")

    return "\n".join(lines).rstrip()


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Read-only report of workload-ranked therapist/time candidates from the real workbook."
        )
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--patient-id", required=True)
    parser.add_argument("--treatment", choices=("ΦΘ", "Ρομποτικό"), default="ΦΘ")
    parser.add_argument("--date", type=_parse_date, required=True)
    args = parser.parse_args()

    workbook = args.workbook.resolve()
    if not workbook.exists():
        print(f"ERROR: workbook not found: {workbook}", file=sys.stderr)
        return 2

    try:
        report = render_candidate_report(
            workbook,
            patient_id=args.patient_id.strip(),
            treatment=args.treatment,
            target_date=args.date,
        )
    except (ValueError, OSError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
