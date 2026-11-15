from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_config.provider_policy_store import load_policy_book  # noqa: E402
from rehab_config.service_hours_store import load_service_hours  # noqa: E402
from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.double_booking_review import build_double_booking_review_items  # noqa: E402
from rehab_core.global_replacement_end_to_end import build_global_replacement_assignments  # noqa: E402
from rehab_core.models import AbsenceKind, Therapist  # noqa: E402
from rehab_core.therapist_absence_queue import build_therapist_absence_replacement_queue  # noqa: E402
from rehab_excel.application_contract import assert_application_contract  # noqa: E402
from rehab_excel.daily_input_reader import read_daily_input  # noqa: E402
from rehab_excel.daily_replacement_facts import (  # noqa: E402
    read_daily_replacement_fact_rows,
    read_daily_replacement_facts,
)
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402
from tools.apply_global_daily_replacements import (  # noqa: E402
    _is_legacy_student_name,
    _read_daily_input_date,
    _replacement_id_factory,
)


def _payload(target_date, result, items):
    return {
        "schema": "eka.double_booking_review.v1",
        "date": target_date.isoformat(),
        "normal_assignments": len(result.assignments),
        "unresolved_sessions": len(result.unresolved_request_ids),
        "requires_user_confirmation": True,
        "items": [
            {
                "target_session_id": item.target_session_id,
                "patient_id": item.patient_id,
                "patient_name": item.patient_name,
                "original_therapist_id": item.original_therapist_id,
                "original_time": item.original_time.strftime("%H:%M"),
                "options": [
                    {
                        "therapist_id": option.therapist_id,
                        "therapist_name": option.therapist_name,
                        "replacement_time": option.replacement_time.strftime("%H:%M"),
                        "warning": option.warning,
                    }
                    for option in item.options
                ],
            }
            for item in items
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build UI-ready double-booking review JSON for an unresolved daily plan."
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_book = args.input.resolve()
    output = args.output.resolve()
    if not input_book.exists():
        print(f"SAFETY STOP: input workbook not found: {input_book}")
        return 2
    if output.exists() and not args.overwrite:
        print(f"SAFETY STOP: output already exists: {output}")
        return 2

    try:
        assert_application_contract(input_book)
        target_date = _read_daily_input_date(input_book)
        patients = read_patient_registry(input_book)
        settings = read_settings(input_book)
        base_entries = read_unified_base_schedule(input_book)
        sessions = materialize_sessions_for_date(base_entries, target_date)
        daily = read_daily_input(
            input_book,
            patients=patients,
            sessions=sessions,
            therapist_names=settings.therapist_names,
        )
        policy_book = load_policy_book(REPO_ROOT / "Config" / "provider_policy.json")
        service_hours = load_service_hours(REPO_ROOT / "Config" / "service_hours.json")
        existing = tuple(
            read_daily_replacement_facts(
                input_book,
                sessions=sessions,
                target_date=target_date,
            )
        )
        history = tuple(
            read_daily_replacement_fact_rows(
                input_book,
                target_date=target_date,
                include_restored=True,
            )
        )
        if not any(a.absence_kind == AbsenceKind.THERAPIST for a in daily.absences):
            raise ValueError("DAILY_INPUT has no therapist absence rows")

        therapists = tuple(
            Therapist(therapist_id=name, display_name=name)
            for name in settings.therapist_names
            if not _is_legacy_student_name(name)
        )
        queue = build_therapist_absence_replacement_queue(
            sessions=sessions,
            absences=daily.absences,
            cancellations=daily.cancellations,
            replacements=existing,
            therapists=therapists,
            patients=patients,
            timeslots=settings.standard_timeslots,
            policy_book=policy_book,
            base_entries=base_entries,
        )
        session_by_id = {session.session_id: session for session in sessions}
        targets = tuple(session_by_id[item.session_id] for item in queue.items)
        result = build_global_replacement_assignments(
            targets,
            therapists=therapists,
            sessions=sessions,
            absences=daily.absences,
            patients=patients,
            replacements=existing,
            cancellations=daily.cancellations,
            timeslots=settings.standard_timeslots,
            policy_book=policy_book,
            base_entries=base_entries,
            service_hours=service_hours,
            replacement_id_factory=_replacement_id_factory(history, session_by_id),
            reason="DAILY_INPUT global therapist absence plan",
        )
        if not result.unresolved_request_ids:
            raise ValueError("Normal global plan has no unresolved sessions")

        items = build_double_booking_review_items(
            result,
            sessions=sessions,
            therapists=therapists,
            patients=patients,
            absences=daily.absences,
            replacements=existing,
            cancellations=daily.cancellations,
            timeslots=settings.standard_timeslots,
            policy_book=policy_book,
            base_entries=base_entries,
        )
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(
            json.dumps(_payload(target_date, result, items), ensure_ascii=False, indent=2),
            encoding="utf-8-sig",
        )
    except (ValueError, KeyError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DOUBLE BOOKING REVIEW PAYLOAD OK")
    print(f"Date: {target_date.isoformat()}")
    print(f"Normal assignments: {len(result.assignments)}")
    print(f"Unresolved sessions: {len(result.unresolved_request_ids)}")
    print(f"Review items: {len(items)}")
    print(f"Output: {output}")
    print("Explicit user confirmation required: True")
    for item in items:
        print(
            f"  {item.patient_name} | {item.original_time.strftime('%H:%M')} | "
            f"options: {len(item.options)}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
