from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.day_patterns import RehabWeekday, parse_day_pattern  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_base_schedule, read_settings  # noqa: E402


DAY_LABELS = {
    RehabWeekday.MONDAY: "Δε",
    RehabWeekday.TUESDAY: "Τρ",
    RehabWeekday.WEDNESDAY: "Τε",
    RehabWeekday.THURSDAY: "Πε",
    RehabWeekday.FRIDAY: "Πα",
}


def _day_text(days) -> str:
    ordered = [day for day in RehabWeekday if day in days]
    return "-".join(DAY_LABELS[day] for day in ordered)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Read-only scan for therapist/time slots where an outpatient can use "
            "at least one weekday without sharing the active cell with an inpatient."
        )
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--therapist")
    parser.add_argument("--limit", type=int, default=20)
    args = parser.parse_args()

    source = args.source.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    patients = read_patient_registry(source)
    patient_by_id = {patient.patient_id: patient for patient in patients}
    entries = read_base_schedule(source)
    settings = read_settings(source)

    occupied: dict[tuple[str, object], set[RehabWeekday]] = {}
    for entry in entries:
        if not entry.therapist_id:
            continue
        patient = patient_by_id.get(entry.patient_id)
        if patient is not None and patient.is_outpatient:
            continue
        occupied.setdefault((entry.therapist_id, entry.start_time), set()).update(
            parse_day_pattern(entry.day_pattern)
        )

    wanted = (args.therapist or "").strip().casefold()
    candidates: list[tuple[int, str, object, set[RehabWeekday]]] = []
    all_days = set(RehabWeekday)
    for therapist in settings.therapist_names:
        if wanted and therapist.casefold() != wanted:
            continue
        for slot in settings.standard_timeslots:
            free = all_days - occupied.get((therapist, slot), set())
            if free:
                candidates.append((len(free), therapist, slot, free))

    candidates.sort(key=lambda item: (-item[0], item[1].casefold(), item[2]))

    print("OUTPATIENT SAFE SLOT SCAN")
    print(f"Source: {source}")
    print(f"Candidate slots: {len(candidates)}")
    for _, therapist, slot, free in candidates[: max(args.limit, 1)]:
        print(f"  {therapist} {slot.strftime('%H:%M')} -> free weekdays: {_day_text(free)}")
    print("Read-only: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
