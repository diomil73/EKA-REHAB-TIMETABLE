from __future__ import annotations

import argparse
import sys
import unicodedata
from datetime import date, datetime, time
from pathlib import Path

HERE = Path(__file__).resolve()
PYTHON_DIR = HERE.parents[1]
REPO_ROOT = HERE.parents[2]
if str(PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(PYTHON_DIR))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.capacity_rebalance import build_capacity_rebalance_cases  # noqa: E402
from rehab_core.models import Therapist  # noqa: E402
from rehab_core.replacement_options import find_replacement_options  # noqa: E402
from rehab_excel.reader import read_base_schedule, read_patients, read_settings  # noqa: E402


def _norm(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().upper())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _is_legacy_student_name(name: str) -> bool:
    compact = _norm(name).replace(" ", "")
    return compact.startswith("ΦΟΙΤ") or compact.startswith("FOIT")


def _parse_date(value: str) -> date:
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    raise argparse.ArgumentTypeError("date must be YYYY-MM-DD or DD/MM/YYYY")


def _fmt_times(slots: tuple[time, ...]) -> str:
    return ", ".join(slot.strftime("%H:%M") for slot in slots)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Read-only planner for providers whose existing base programme is over "
            "daily timeslot capacity. No workbook is modified."
        )
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--date", type=_parse_date, default=date(2026, 9, 18))
    parser.add_argument("--provider", default=None)
    parser.add_argument("--limit", type=int, default=3)
    args = parser.parse_args()

    workbook = args.workbook.resolve()
    if not workbook.exists():
        print(f"SAFETY STOP: workbook not found: {workbook}")
        return 2

    patients = read_patients(workbook)
    patient_map = {patient.patient_id: patient for patient in patients}
    settings = read_settings(workbook)
    entries = read_base_schedule(workbook)
    sessions = materialize_sessions_for_date(entries, args.date)
    therapists = tuple(
        Therapist(name, name)
        for name in settings.therapist_names
        if not _is_legacy_student_name(name)
    )

    cases = build_capacity_rebalance_cases(
        target_date=args.date,
        sessions=sessions,
        therapists=therapists,
    )
    if args.provider:
        wanted = _norm(args.provider)
        cases = tuple(case for case in cases if _norm(case.provider_id) == wanted)

    print("CAPACITY REBALANCE PLAN")
    print(f"Date: {args.date.isoformat()}")
    print(f"Workbook: {workbook}")
    print("READ-ONLY: no Excel file is modified.")

    if not cases:
        print("OVER CAPACITY: NONE for the selected date/provider.")
        return 0

    therapist_ids = {therapist.therapist_id for therapist in therapists}
    for case in cases:
        print()
        print(
            f"{case.provider_id}: {case.used_timeslots}/{case.max_timeslots} "
            f"timeslots | must free at least {case.excess_timeslots} slot(s)"
        )
        print(
            "IMPORTANT: if a clock time contains multiple active patients, ALL of "
            "them must move before that timeslot is actually freed."
        )

        for slot_index, slot in enumerate(case.slots, start=1):
            patient_names = [
                patient_map.get(session.patient_id).display_name
                if patient_map.get(session.patient_id) is not None
                else session.patient_id
                for session in slot.sessions
            ]
            print(
                f"  SLOT {slot_index}. {slot.start_time.strftime('%H:%M')} | "
                + " + ".join(patient_names)
            )

            for session in slot.sessions:
                patient = patient_map.get(session.patient_id)
                patient_name = patient.display_name if patient is not None else session.patient_id
                options = find_replacement_options(
                    target_session=session,
                    therapists=therapists,
                    sessions=sessions,
                    patients=patients,
                    requested_time=session.start_time,
                    timeslots=settings.standard_timeslots,
                )
                # Defensive filter in case a future provider registry contains aliases.
                options = [
                    option
                    for option in options
                    if option.provider_id in therapist_ids
                    and option.provider_id != case.provider_id
                ]
                print(f"     {patient_name}:")
                if not options:
                    print("       no feasible replacement option")
                    continue
                for rank, option in enumerate(options[: max(args.limit, 1)], start=1):
                    exact = "exact" if option.exact_time_available else "alternative"
                    capacity = ""
                    if option.capacity_limit is not None and option.capacity_remaining is not None:
                        used = option.capacity_limit - option.capacity_remaining
                        capacity = f" | slots {used}/{option.capacity_limit}"
                    print(
                        f"       {rank}. {option.display_name} | load {option.active_sessions}"
                        f"{capacity} | {option.recommended_time.strftime('%H:%M')} ({exact})"
                        f" | free: {_fmt_times(option.available_timeslots)}"
                    )

    print()
    print(
        "RULE: this tool proposes possibilities only. It never chooses which existing "
        "patient to move and never rewrites the base programme automatically."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
