from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.permanent_assignment import check_permanent_assignment
from rehab_excel.permanent_preview import (
    PermanentPreviewError,
    create_permanent_assignment_preview,
    parse_hhmm,
)
from rehab_excel.reader import read_base_schedule


def _find_entry(entries, patient_name, from_provider, from_time):
    return tuple(
        e for e in entries
        if e.therapist_id == from_provider
        and e.start_time == from_time
        and e.treatment == "ΦΘ"
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a safe permanent-assignment preview copy of the rehab workbook."
    )
    parser.add_argument("--patient", required=True)
    parser.add_argument("--from-provider", required=True)
    parser.add_argument("--from-time", required=True)
    parser.add_argument("--to-provider", required=True)
    parser.add_argument("--to-time", required=True)
    parser.add_argument(
        "--source",
        default=str(REPO_ROOT / "Rehab_Center_System_v27_1.xlsm"),
    )
    parser.add_argument(
        "--output",
        default=str(REPO_ROOT / "Excel" / "previews" / "PERMANENT_ASSIGNMENT_PREVIEW.xlsm"),
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        from_time = parse_hhmm(args.from_time)
        to_time = parse_hhmm(args.to_time)

        from rehab_excel.permanent_preview import locate_physio_assignment
        located = locate_physio_assignment(
            args.source,
            patient_name=args.patient,
            from_provider=args.from_provider,
            from_time=from_time,
        )

        entries = tuple(read_base_schedule(args.source))
        source_entry_id = f"planner:{located.row}:ΦΘ"
        source_entry = next(
            (entry for entry in entries if entry.base_entry_id == source_entry_id),
            None,
        )
        if source_entry is None:
            raise PermanentPreviewError(
                f"Could not resolve {source_entry_id} in imported base schedule"
            )

        check = check_permanent_assignment(
            provider_id=args.to_provider,
            max_daily_timeslots=6,
            existing_entries=entries,
            proposed_day_pattern=source_entry.day_pattern,
            proposed_time=to_time,
            source_entry_id=source_entry.base_entry_id,
            patient_id=source_entry.patient_id,
        )
        if not check.allowed:
            print("SAFETY STOP: proposed permanent assignment is not allowed.")
            entry_map = {entry.base_entry_id: entry for entry in entries}
            for day in check.days:
                status = []
                if day.over_capacity:
                    status.append(
                        f"capacity {day.used_timeslots_before}/6 -> {day.used_timeslots_after}/6"
                    )
                if day.has_conflict:
                    status.append("provider same-day conflict")
                if day.has_patient_conflict:
                    therapies = []
                    for entry_id in day.patient_conflicting_entry_ids:
                        conflict = entry_map.get(entry_id)
                        therapies.append(
                            entry_id
                            if conflict is None
                            else f"{conflict.treatment} {conflict.start_time.strftime('%H:%M')}"
                        )
                    status.append("patient conflict: " + ", ".join(therapies))
                if status:
                    print(f"  {day.weekday.name}: " + ", ".join(status))
            return 2

        report = create_permanent_assignment_preview(
            source_path=args.source,
            output_path=args.output,
            patient_name=args.patient,
            from_provider=args.from_provider,
            from_time=from_time,
            to_provider=args.to_provider,
            to_time=to_time,
            overwrite=args.overwrite,
        )
        print("PERMANENT ASSIGNMENT PREVIEW OK")
        print(f"Source: {report.source_path}")
        print(f"Preview: {report.output_path}")
        print(
            f"Changed PATIENT_PLANNER row {report.row}: "
            f"{report.patient_name} | "
            f"{report.old_provider} {report.old_time} -> "
            f"{report.new_provider} {report.new_time} [{report.day_pattern}]"
        )
        print(f"Source unchanged: {report.source_unchanged}")
        print(f"VBA preserved: {report.vba_preserved}")
        print(
            "NEXT: open ONLY the preview workbook. In Excel, use the existing "
            "ΑΝΑΝΕΩΣΗ/Refresh action to rebuild the operational views, then inspect "
            "the moved patient and provider capacity."
        )
        return 0
    except PermanentPreviewError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
