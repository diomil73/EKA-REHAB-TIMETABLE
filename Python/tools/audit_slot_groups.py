from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.day_patterns import RehabWeekday  # noqa: E402
from rehab_core.slot_group_audit import audit_slot_groups, patient_name_map  # noqa: E402
from rehab_excel.reader import read_base_schedule, read_patients  # noqa: E402


_DAY_LABEL = {
    RehabWeekday.MONDAY: "Δε",
    RehabWeekday.TUESDAY: "Τρ",
    RehabWeekday.WEDNESDAY: "Τε",
    RehabWeekday.THURSDAY: "Πε",
    RehabWeekday.FRIDAY: "Πα",
}


def _day_list(days) -> str:
    return "-".join(_DAY_LABEL[day] for day in sorted(days))


def _resolve_workbook(value: str | None) -> Path:
    if value:
        path = Path(value).expanduser().resolve()
    else:
        path = ROOT / "Rehab_Center_System_v27_1.xlsm"
    if not path.exists():
        raise SystemExit(f"Workbook not found: {path}")
    return path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit patient-centric therapist/time slot grouping in the base workbook."
    )
    parser.add_argument("--workbook", help="Path to .xlsm workbook; defaults to v27.1 in repo root")
    args = parser.parse_args()

    workbook = _resolve_workbook(args.workbook)
    entries = read_base_schedule(workbook)
    patients = read_patients(workbook)
    names = patient_name_map(patients)
    report = audit_slot_groups(entries)

    print("PATIENT-CENTRIC SLOT GROUP AUDIT OK")
    print(f"Workbook: {workbook}")
    print(f"Base assignments: {report.base_assignment_count}")
    print(f"Therapist/time groups: {report.group_count}")
    print(f"Single-member groups: {report.single_member_group_count}")
    print(f"Multi-member groups: {report.multi_member_group_count}")
    print(f"Clean derived pairs: {report.clean_pair_count}")
    print(f"Overlapping-day groups requiring review: {report.overlapping_group_count}")

    pairs = [group for group in report.groups if group.is_pair]
    if pairs:
        print("\nDERIVED PAIRS")
        for group in pairs:
            members = " + ".join(
                f"{names.get(member.patient_id, member.patient_id)} [{member.day_pattern}]"
                for member in group.members
            )
            print(f"  {group.therapist_id} {group.start_time.strftime('%H:%M')} | {members}")

    overlaps = [group for group in report.groups if group.overlaps]
    if overlaps:
        print("\nOVERLAPPING-DAY GROUPS (REVIEW, DO NOT AUTO-MERGE AS A CLEAN PAIR)")
        for group in overlaps:
            members = " + ".join(
                f"{names.get(member.patient_id, member.patient_id)} [{member.day_pattern}]"
                for member in group.members
            )
            shared = sorted({day for overlap in group.overlaps for day in overlap.weekdays})
            print(
                f"  {group.therapist_id} {group.start_time.strftime('%H:%M')} | "
                f"{members} | shared days: {_day_list(shared)}"
            )

    print("\nRULE: each patient assignment remains independent; pairs/groups are derived views only.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
