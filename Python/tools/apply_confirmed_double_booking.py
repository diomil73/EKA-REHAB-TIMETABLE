from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime, time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_config.provider_policy_store import load_policy_book  # noqa: E402
from rehab_config.service_hours_store import load_service_hours  # noqa: E402
from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.double_booking_exception import (  # noqa: E402
    confirm_double_booking_exception,
    propose_double_booking_exception,
)
from rehab_core.global_replacement_end_to_end import build_global_replacement_assignments  # noqa: E402
from rehab_core.models import AbsenceKind, Therapist  # noqa: E402
from rehab_core.therapist_absence_queue import build_therapist_absence_replacement_queue  # noqa: E402
from rehab_excel.application_contract import assert_application_contract  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_reader import read_daily_input  # noqa: E402
from rehab_excel.daily_replacement_facts import (  # noqa: E402
    append_daily_replacement_facts,
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


EXPECTED_SCHEMA = "eka.double_booking_review.v1"


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Time must be HH:MM") from exc


def _load_review_choice(
    review_path: Path,
    *,
    target_session_id: str,
    therapist_id: str,
    replacement_time: time,
) -> dict:
    payload = json.loads(review_path.read_text(encoding="utf-8-sig"))
    if payload.get("schema") != EXPECTED_SCHEMA:
        raise ValueError("Unsupported double-booking review schema")
    if payload.get("requires_user_confirmation") is not True:
        raise ValueError("Review payload does not require explicit confirmation")

    for item in payload.get("items", []):
        if str(item.get("target_session_id", "")) != target_session_id:
            continue
        for option in item.get("options", []):
            if (
                str(option.get("therapist_id", "")) == therapist_id
                and str(option.get("replacement_time", ""))
                == replacement_time.strftime("%H:%M")
            ):
                return option
        raise ValueError("Selected therapist/time is not present in review options")
    raise ValueError("Selected target session is not present in review payload")


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Apply one explicitly confirmed double-booking exception to a fresh preview copy."
        )
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--review", required=True, type=Path)
    parser.add_argument("--target-session-id", required=True)
    parser.add_argument("--therapist-id", required=True)
    parser.add_argument("--time", required=True, type=_parse_time)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--confirm", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    if not args.confirm:
        print("SAFETY STOP: explicit --confirm is required for double booking")
        return 2

    input_book = args.input.resolve()
    review_path = args.review.resolve()
    output = args.output.resolve()
    if not input_book.exists():
        print(f"SAFETY STOP: input workbook not found: {input_book}")
        return 2
    if not review_path.exists():
        print(f"SAFETY STOP: review payload not found: {review_path}")
        return 2
    if output.exists() and not args.overwrite:
        print(f"SAFETY STOP: output already exists: {output}")
        return 2

    source_hash = file_sha256(input_book)
    output_created = False
    try:
        _load_review_choice(
            review_path,
            target_session_id=args.target_session_id,
            therapist_id=args.therapist_id,
            replacement_time=args.time,
        )
        input_contract = assert_application_contract(input_book)
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
        normal_result = build_global_replacement_assignments(
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
        if args.target_session_id not in normal_result.unresolved_request_ids:
            raise ValueError(
                "Selected target is no longer unresolved by the normal global plan"
            )
        target = session_by_id.get(args.target_session_id)
        if target is None:
            raise ValueError("Unknown target session")

        proposal = propose_double_booking_exception(
            target_session=target,
            unresolved_session_ids=normal_result.unresolved_request_ids,
            therapist_id=args.therapist_id,
            replacement_time=args.time,
            therapists=therapists,
            sessions=sessions,
            absences=daily.absences,
            patients=patients,
            replacements=existing + normal_result.assignments,
            cancellations=daily.cancellations,
            policy_book=policy_book,
            base_entries=base_entries,
        )
        replacement_id = _replacement_id_factory(history, session_by_id)(
            target.session_id
        )
        assignment = confirm_double_booking_exception(
            proposal,
            target_session=target,
            replacement_id=replacement_id,
            confirmed=True,
        )

        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists():
            output.unlink()
        shutil.copy2(input_book, output)
        output_created = True
        append_daily_replacement_facts(output, (assignment,), sessions=sessions)
        output_contract = assert_application_contract(output)

        reread = read_daily_replacement_facts(
            output,
            sessions=sessions,
            target_date=target_date,
        )
        matched = [
            item
            for item in reread
            if item.replacement_id == assignment.replacement_id
        ]
        if len(matched) != 1:
            raise ValueError("Confirmed double-booking fact did not read back exactly once")
    except Exception as exc:
        if output_created and output.exists():
            try:
                output.unlink()
            except OSError:
                pass
        print(f"SAFETY STOP: {exc}")
        return 2

    source_unchanged = file_sha256(input_book) == source_hash
    print("CONFIRMED DOUBLE BOOKING PREVIEW OK")
    print(f"Date: {target_date.isoformat()}")
    print(f"Target session: {assignment.target_session_id}")
    print(f"Patient: {assignment.patient_id}")
    print(f"Therapist: {assignment.replacement_therapist_id}")
    print(f"Time: {assignment.replacement_time.strftime('%H:%M')}")
    print(f"Reason: {assignment.reason}")
    print(f"Preview: {output}")
    print(f"Input unchanged: {source_unchanged}")
    print("Application contract input: OK")
    print("Application contract output: OK")
    print(f"Output VBA component count: {len(output_contract.component_names)}")
    return 0 if source_unchanged else 3


if __name__ == "__main__":
    raise SystemExit(main())
