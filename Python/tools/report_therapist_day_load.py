from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import date
from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.reader import read_base_schedule, read_patients  # noqa: E402
from rehab_excel.workbook_scheduling_candidates import (  # noqa: E402
    materialize_recurring_sessions,
)


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def render_therapist_day_load_report(
    workbook: Path,
    *,
    therapist: str,
    target_date: date,
) -> str:
    entries = read_base_schedule(workbook)
    sessions = materialize_recurring_sessions(entries, target_date)
    patients = {patient.patient_id: patient.display_name for patient in read_patients(workbook)}

    therapist_key = therapist.strip().casefold()
    selected = [
        session
        for session in sessions
        if session.therapist_id.strip().casefold() == therapist_key
    ]

    by_time: dict[object, list[object]] = defaultdict(list)
    for session in selected:
        by_time[session.start_time].append(session)

    lines = [
        "THERAPIST DAY LOAD",
        f"Workbook: {workbook}",
        f"Therapist: {therapist}",
        f"Date: {target_date.isoformat()}",
        f"Sessions: {len(selected)}",
        f"Occupied slots: {len(by_time)}",
        "",
    ]

    if not selected:
        lines.append("No sessions found for this therapist/date.")
        return "\n".join(lines)

    for slot in sorted(by_time):
        slot_sessions = by_time[slot]
        marker = "  <-- SHARED SLOT" if len(slot_sessions) > 1 else ""
        lines.append(f"{slot.strftime('%H:%M')} ({len(slot_sessions)} session{'s' if len(slot_sessions) != 1 else ''}){marker}")
        for session in slot_sessions:
            patient_name = patients.get(session.patient_id, "?")
            treatment = session.treatment or "?"
            lines.append(
                f"   - PatientID {session.patient_id} | {patient_name} | {treatment} | {session.session_id}"
            )
        lines.append("")

    return "\n".join(lines).rstrip()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read-only diagnostic of one therapist's real recurring load for one date."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--therapist", required=True)
    parser.add_argument("--date", type=_parse_date, required=True)
    args = parser.parse_args()

    workbook = args.workbook.resolve()
    if not workbook.exists():
        print(f"ERROR: workbook not found: {workbook}", file=sys.stderr)
        return 2

    try:
        report = render_therapist_day_load_report(
            workbook,
            therapist=args.therapist.strip(),
            target_date=args.date,
        )
    except (ValueError, OSError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
