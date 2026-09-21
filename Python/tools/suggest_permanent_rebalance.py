from __future__ import annotations

import argparse
import sys
import unicodedata
from collections import defaultdict
from datetime import time
from pathlib import Path

HERE = Path(__file__).resolve()
PYTHON_DIR = HERE.parents[1]
REPO_ROOT = HERE.parents[2]
if str(PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(PYTHON_DIR))

from rehab_core.day_patterns import RehabWeekday  # noqa: E402
from rehab_core.permanent_rebalance import (  # noqa: E402
    find_permanent_rebalance_options,
    provider_weekly_capacity,
)
from rehab_excel.reader import read_base_schedule, read_patients, read_settings  # noqa: E402


DAY_LABEL = {
    RehabWeekday.MONDAY: "Δε",
    RehabWeekday.TUESDAY: "Τρ",
    RehabWeekday.WEDNESDAY: "Τε",
    RehabWeekday.THURSDAY: "Πε",
    RehabWeekday.FRIDAY: "Πα",
}


def _norm(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().upper())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _is_legacy_student_name(name: str) -> bool:
    compact = _norm(name).replace(" ", "")
    return compact.startswith("ΦΟΙΤ") or compact.startswith("FOIT")


def _days(days) -> str:
    return "-".join(DAY_LABEL[day] for day in days) if days else "-"


def _projection(check) -> str:
    parts = []
    for day in check.days:
        parts.append(
            f"{DAY_LABEL[day.weekday]} {day.used_timeslots_before}/{day.max_timeslots}"
            f"→{day.used_timeslots_after}/{day.max_timeslots}"
        )
    return ", ".join(parts)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Read-only permanent rebalance suggestions. Shows only moves that "
            "reduce source over-capacity and pass weekly destination capacity/conflict checks."
        )
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--provider", required=True)
    parser.add_argument("--per-patient", type=int, default=3)
    args = parser.parse_args()

    workbook = args.workbook.resolve()
    if not workbook.exists():
        print(f"SAFETY STOP: workbook not found: {workbook}")
        return 2

    patients = read_patients(workbook)
    patient_map = {patient.patient_id: patient for patient in patients}
    settings = read_settings(workbook)
    entries = read_base_schedule(workbook)

    provider_lookup = {_norm(name): name for name in settings.therapist_names}
    provider = provider_lookup.get(_norm(args.provider))
    if provider is None:
        print(f"SAFETY STOP: provider not found in SETTINGS: {args.provider}")
        return 2

    therapist_names = tuple(
        name for name in settings.therapist_names if not _is_legacy_student_name(name)
    )
    limits = {name: 6 for name in therapist_names}

    before = provider_weekly_capacity(
        provider_id=provider,
        existing_entries=entries,
        max_daily_timeslots=6,
    )
    if not any(item.excess > 0 for item in before):
        print("PERMANENT REBALANCE SUGGESTIONS")
        print(f"Workbook: {workbook}")
        print(f"Provider: {provider}")
        print("OVER CAPACITY: NO. No permanent rebalance is required.")
        return 0

    options = find_permanent_rebalance_options(
        source_provider_id=provider,
        existing_entries=entries,
        destination_provider_ids=therapist_names,
        standard_timeslots=settings.standard_timeslots,
        provider_capacity_limits=limits,
        source_capacity_limit=6,
    )

    grouped = defaultdict(list)
    source_order = []
    for option in options:
        key = option.source_entry.base_entry_id
        if key not in grouped:
            source_order.append(key)
        grouped[key].append(option)

    print("PERMANENT REBALANCE SUGGESTIONS")
    print(f"Workbook: {workbook}")
    print(f"Provider: {provider}")
    print("READ-ONLY: no Excel file is modified.")
    print(
        "Current weekly capacity: "
        + ", ".join(
            f"{DAY_LABEL[item.weekday]} {item.used_timeslots}/{item.max_timeslots}"
            for item in before
        )
    )
    print(
        "RULE: only moves that reduce source over-capacity AND are safe across the "
        "patient's full recurring day-pattern are shown."
    )

    if not options:
        print("NO FEASIBLE PERMANENT MOVES found under current weekly capacity/conflict rules.")
        return 0

    max_per_patient = max(1, args.per_patient)
    shown = 0
    for key in source_order:
        patient_options = grouped[key]
        first = patient_options[0]
        patient = patient_map.get(first.source_entry.patient_id)
        patient_name = patient.display_name if patient is not None else first.source_entry.patient_id
        print()
        print(
            f"{patient_name} | {provider} {first.source_entry.start_time.strftime('%H:%M')} "
            f"[{first.source_entry.day_pattern}] | frees: {_days(first.source_days_freed)}"
        )
        print(
            f"  source excess: {first.source_excess_before} -> {first.source_excess_after} | "
            f"over-capacity days: {first.source_overcapacity_days_before} -> "
            f"{first.source_overcapacity_days_after}"
        )
        for rank, option in enumerate(patient_options[:max_per_patient], start=1):
            exact = "same time" if option.same_time else "new time"
            print(
                f"  {rank}. {option.destination_provider_id} "
                f"{option.destination_time.strftime('%H:%M')} ({exact}) | "
                f"destination: {_projection(option.destination_check)}"
            )
        shown += 1

    print()
    print(f"Patients/assignments with at least one feasible permanent move: {shown}")
    print(
        "NOTE: ranking is a planning aid only. It does not choose which patient should "
        "move and does not rewrite PATIENT_PLANNER."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
