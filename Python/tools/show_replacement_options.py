from __future__ import annotations

import argparse
import sys
import unicodedata
from datetime import date, datetime, time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.models import Therapist  # noqa: E402
from rehab_core.replacement_options import find_replacement_options  # noqa: E402
from rehab_excel.reader import (  # noqa: E402
    read_base_schedule,
    read_patients,
    read_settings,
)


def _norm(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().upper())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Time must be HH:MM") from exc


def _is_legacy_student_name(name: str) -> bool:
    norm = _norm(name).replace(" ", "")
    return norm.startswith("ΦΟΙΤ") or norm.startswith("FOIT")


def _find_target_session(sessions, patients, patient_name: str, target_time: time):
    wanted = _norm(patient_name)
    patient_matches = [p for p in patients if _norm(p.display_name) == wanted]
    if len(patient_matches) != 1:
        raise ValueError(
            f"Expected one patient named {patient_name!r}; found {len(patient_matches)}"
        )
    patient = patient_matches[0]
    matches = [
        s
        for s in sessions
        if s.patient_id == patient.patient_id and s.start_time == target_time
    ]
    if len(matches) != 1:
        raise ValueError(
            f"Expected one active session for {patient.display_name} at "
            f"{target_time.strftime('%H:%M')}; found {len(matches)}"
        )
    return patient, matches[0]


def _fmt_times(slots: tuple[time, ...]) -> str:
    return ", ".join(slot.strftime("%H:%M") for slot in slots)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Show ranked replacement providers and their feasible alternative "
            "timeslots for one real patient session. Read-only: no Excel write-back."
        )
    )
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument("--patient", required=True)
    parser.add_argument("--time", type=_parse_time, required=True)
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--limit", type=int, default=10)
    args = parser.parse_args()

    source = args.source.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    try:
        patients = read_patients(source)
        settings = read_settings(source)
        base_entries = read_base_schedule(source)
        sessions = materialize_sessions_for_date(base_entries, args.date)
        patient, target = _find_target_session(
            sessions, patients, args.patient, args.time
        )
        if target.robotic:
            print(
                "SAFETY STOP: robotic capability is not yet authoritative in "
                "the workbook adapter; do not rank robotic replacements from this tool."
            )
            return 2

        therapists = [
            Therapist(therapist_id=name, display_name=name)
            for name in settings.therapist_names
            if not _is_legacy_student_name(name)
        ]
        timeslots = settings.standard_timeslots
        options = find_replacement_options(
            target_session=target,
            therapists=therapists,
            sessions=sessions,
            patients=patients,
            requested_time=args.time,
            timeslots=timeslots,
        )
    except ValueError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("REPLACEMENT OPTIONS OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Patient: {patient.display_name}")
    print(f"Original: {target.therapist_id} {target.start_time.strftime('%H:%M')}")
    print(f"Timeslot grid: {_fmt_times(tuple(timeslots))}")
    print(f"Candidates: {len(options)}")
    if not options:
        print("  No feasible replacement providers/timeslots found.")
        return 0

    for rank, option in enumerate(options[: max(args.limit, 0)], start=1):
        exact = "exact" if option.exact_time_available else "alternative"
        print(
            f"  {rank}. {option.display_name} | load {option.active_sessions} | "
            f"recommended {option.recommended_time.strftime('%H:%M')} ({exact}) | "
            f"free: {_fmt_times(option.available_timeslots)}"
        )
    print(
        "RULE: provider ranking stays load-first; a busy requested time lowers "
        "time-match priority but does not exclude a provider when another common slot exists."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
