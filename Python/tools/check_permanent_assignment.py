from __future__ import annotations

import argparse
import sys
import unicodedata
from datetime import datetime, time
from pathlib import Path

HERE = Path(__file__).resolve()
PYTHON_DIR = HERE.parents[1]
REPO_ROOT = HERE.parents[2]
if str(PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(PYTHON_DIR))

from rehab_core.day_patterns import RehabWeekday  # noqa: E402
from rehab_core.permanent_assignment import check_permanent_assignment  # noqa: E402
from rehab_excel.reader import read_base_schedule, read_patients, read_settings  # noqa: E402

DAY_LABELS = {
    RehabWeekday.MONDAY: "Δε",
    RehabWeekday.TUESDAY: "Τρ",
    RehabWeekday.WEDNESDAY: "Τε",
    RehabWeekday.THURSDAY: "Πε",
    RehabWeekday.FRIDAY: "Πα",
}


def _norm(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().casefold())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("time must be HH:MM") from exc


def _provider_limit(provider_name: str) -> int:
    compact = _norm(provider_name).replace(" ", "")
    return 5 if compact.startswith("φοιτ") or compact.startswith("foit") else 6


def _resolve_provider(requested: str, provider_names: tuple[str, ...]) -> str | None:
    wanted = _norm(requested)
    exact = [name for name in provider_names if _norm(name) == wanted]
    return exact[0] if len(exact) == 1 else None


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Read-only weekly provider capacity and patient cross-specialty conflict "
            "check before a permanent therapist/time assignment."
        )
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--patient", required=True)
    parser.add_argument("--from-provider", default=None)
    parser.add_argument("--from-time", type=_parse_time, default=None)
    parser.add_argument("--to-provider", required=True)
    parser.add_argument("--to-time", required=True, type=_parse_time)
    parser.add_argument(
        "--to-days",
        default=None,
        help="Optional new recurring day-pattern; defaults to the source assignment pattern.",
    )
    args = parser.parse_args()

    workbook = args.workbook.resolve()
    if not workbook.exists():
        print(f"SAFETY STOP: workbook not found: {workbook}")
        return 2

    patients = read_patients(workbook)
    entries = read_base_schedule(workbook)
    settings = read_settings(workbook)

    patient_matches = [p for p in patients if _norm(p.display_name) == _norm(args.patient)]
    if not patient_matches:
        print(f"SAFETY STOP: patient not found: {args.patient}")
        return 2
    patient_ids = {p.patient_id for p in patient_matches}

    candidates = [entry for entry in entries if entry.patient_id in patient_ids]
    if args.from_provider:
        candidates = [
            entry for entry in candidates
            if entry.therapist_id and _norm(entry.therapist_id) == _norm(args.from_provider)
        ]
    if args.from_time:
        candidates = [entry for entry in candidates if entry.start_time == args.from_time]

    if not candidates:
        print("SAFETY STOP: no base assignment matched the patient/source filters.")
        return 2
    if len(candidates) > 1:
        print("SAFETY STOP: multiple base assignments match. Add --from-provider and/or --from-time.")
        for entry in candidates:
            print(
                f"  {entry.base_entry_id} | {entry.therapist_id} "
                f"{entry.start_time.strftime('%H:%M')} | {entry.day_pattern} | {entry.treatment}"
            )
        return 2

    source = candidates[0]
    target_provider = _resolve_provider(args.to_provider, tuple(settings.therapist_names))
    if target_provider is None:
        print(f"SAFETY STOP: destination provider not found uniquely: {args.to_provider}")
        return 2

    proposed_days = args.to_days or source.day_pattern
    limit = _provider_limit(target_provider)
    result = check_permanent_assignment(
        provider_id=target_provider,
        max_daily_timeslots=limit,
        existing_entries=entries,
        proposed_day_pattern=proposed_days,
        proposed_time=args.to_time,
        source_entry_id=source.base_entry_id,
        patient_id=source.patient_id,
    )

    patient_name = patient_matches[0].display_name
    entry_map = {entry.base_entry_id: entry for entry in entries}
    patient_name_by_id = {p.patient_id: p.display_name for p in patients}

    print("PERMANENT ASSIGNMENT CAPACITY / PATIENT-SCHEDULE CHECK")
    print(f"Workbook: {workbook}")
    print(f"Patient: {patient_name} (id={source.patient_id})")
    print(
        f"Current: {source.therapist_id} {source.start_time.strftime('%H:%M')} "
        f"[{source.day_pattern}]"
    )
    print(
        f"Proposed: {target_provider} {args.to_time.strftime('%H:%M')} "
        f"[{proposed_days}] | capacity limit {limit}"
    )
    print("READ-ONLY: no Excel file is modified.")
    print()
    print("WEEKLY PROJECTION")
    for day in result.days:
        status_parts = []
        if day.over_capacity:
            status_parts.append("OVER CAPACITY")
        if day.has_conflict:
            names = []
            for entry_id in day.conflicting_entry_ids:
                conflict_entry = entry_map.get(entry_id)
                if conflict_entry is None:
                    names.append(entry_id)
                else:
                    names.append(
                        patient_name_by_id.get(conflict_entry.patient_id, conflict_entry.patient_id)
                    )
            status_parts.append("PROVIDER CONFLICT: " + ", ".join(names))
        if day.has_patient_conflict:
            therapies = []
            for entry_id in day.patient_conflicting_entry_ids:
                conflict_entry = entry_map.get(entry_id)
                if conflict_entry is None:
                    therapies.append(entry_id)
                else:
                    therapies.append(
                        f"{conflict_entry.treatment} {conflict_entry.start_time.strftime('%H:%M')}"
                    )
            status_parts.append("PATIENT CONFLICT: " + ", ".join(therapies))
        status = "OK" if not status_parts else " | ".join(status_parts)
        print(
            f"  {DAY_LABELS[day.weekday]}: "
            f"{day.used_timeslots_before}/{day.max_timeslots} -> "
            f"{day.used_timeslots_after}/{day.max_timeslots} | {status}"
        )

    print()
    if result.allowed:
        print("RESULT: ALLOWED by provider capacity and patient-schedule rules.")
        print("NOTE: this is validation only; the base programme was not changed.")
        return 0

    print(
        "RESULT: BLOCKED. Permanent assignment would violate provider capacity/time "
        "or overlap another treatment of the same patient."
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
