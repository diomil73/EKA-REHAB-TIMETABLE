from __future__ import annotations

import argparse
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve()
PYTHON_DIR = HERE.parents[1]
REPO_ROOT = HERE.parents[2]
if str(PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(PYTHON_DIR))

from rehab_core.day_patterns import RehabWeekday, parse_day_pattern  # noqa: E402
from rehab_excel.reader import read_base_schedule  # noqa: E402

DAY_LABELS = {
    RehabWeekday.MONDAY: "Δε",
    RehabWeekday.TUESDAY: "Τρ",
    RehabWeekday.WEDNESDAY: "Τε",
    RehabWeekday.THURSDAY: "Πε",
    RehabWeekday.FRIDAY: "Πα",
}


def _norm(text: str) -> str:
    value = unicodedata.normalize("NFD", text.casefold())
    return "".join(ch for ch in value if unicodedata.category(ch) != "Mn").strip()


def capacity_limit(provider_name: str) -> int:
    """Legacy-safe capacity lookup until the real student registry is wired in."""
    name = _norm(provider_name)
    if name.startswith("φοιτ") or name.startswith("student"):
        return 5
    return 6


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit recurring provider timeslot capacity from PATIENT_PLANNER."
    )
    parser.add_argument(
        "--workbook",
        default=str(REPO_ROOT / "Rehab_Center_System_v27_1.xlsm"),
        help="Path to the source .xlsm workbook",
    )
    args = parser.parse_args()

    workbook = Path(args.workbook).resolve()
    if not workbook.exists():
        print(f"SAFETY STOP: workbook not found: {workbook}")
        return 2

    entries = read_base_schedule(workbook)
    occupied: dict[str, dict[RehabWeekday, set]] = defaultdict(
        lambda: defaultdict(set)
    )
    session_counts: dict[str, dict[RehabWeekday, int]] = defaultdict(
        lambda: defaultdict(int)
    )

    for entry in entries:
        if not entry.therapist_id:
            continue
        for weekday in parse_day_pattern(entry.day_pattern):
            occupied[entry.therapist_id][weekday].add(entry.start_time)
            session_counts[entry.therapist_id][weekday] += 1

    over = []
    full = []
    normal = []
    for provider in sorted(occupied, key=str.casefold):
        limit = capacity_limit(provider)
        day_rows = []
        max_used = 0
        for weekday in RehabWeekday:
            used = len(occupied[provider].get(weekday, set()))
            sessions = session_counts[provider].get(weekday, 0)
            max_used = max(max_used, used)
            day_rows.append((weekday, used, sessions))
        record = (provider, limit, max_used, day_rows)
        if max_used > limit:
            over.append(record)
        elif max_used == limit:
            full.append(record)
        else:
            normal.append(record)

    print("PROVIDER CAPACITY AUDIT")
    print(f"Workbook: {workbook}")
    print(f"Imported base assignments: {len(entries)}")
    print(
        "Rule: blank day-pattern + valid time = Καθ/να "
        "(Δε-Τρ-Τε-Πε-Πα)."
    )
    print("Capacity: physiotherapist 6 timeslots/day; student 5 timeslots/day.")
    print("Pairs sharing the same clock time count as one occupied timeslot.")
    print()

    def emit(title: str, records) -> None:
        print(title)
        if not records:
            print("  NONE")
            return
        for provider, limit, max_used, day_rows in records:
            days = ", ".join(
                f"{DAY_LABELS[d]} {used}/{limit}"
                + (f" ({sessions} patients)" if sessions != used else "")
                for d, used, sessions in day_rows
                if used
            )
            print(f"  {provider} | max {max_used}/{limit} | {days}")

    emit("OVER CAPACITY", over)
    print()
    emit("AT CAPACITY", full)
    print()
    print(
        "RESULT: "
        + (f"{len(over)} provider(s) exceed capacity." if over else "no provider exceeds capacity.")
    )
    return 1 if over else 0


if __name__ == "__main__":
    raise SystemExit(main())
