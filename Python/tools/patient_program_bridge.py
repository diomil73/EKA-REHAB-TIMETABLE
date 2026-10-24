from __future__ import annotations

import argparse
from datetime import date, datetime, time
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.day_patterns import parse_day_pattern
from rehab_excel.automatic_patient_schedule import (
    AutomaticTreatmentRequest,
    propose_automatic_patient_schedule,
)
from rehab_excel.patient_program_draft import build_patient_program_draft
from rehab_excel.reader import read_base_schedule, read_settings


STANDARD_TREATMENT_DURATION_MINUTES = 45
RECLINED_TREATMENT_DURATION_MINUTES = 30
_WEEKDAY_NAMES = {
    0: "Δευτέρα",
    1: "Τρίτη",
    2: "Τετάρτη",
    3: "Πέμπτη",
    4: "Παρασκευή",
}


def _safe(value: object) -> str:
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def _emit_config(workbook: Path, patient_id: str) -> None:
    draft = build_patient_program_draft(workbook, patient_id=patient_id)
    settings = read_settings(workbook)

    print("\t".join(("OK", _safe(draft.patient_id))))
    for index, spec in enumerate(draft.treatments, start=1):
        print(
            "\t".join(
                (
                    "TREATMENT",
                    str(index),
                    _safe(spec.treatment),
                    _safe(spec.provider_kind),
                    "1" if spec.requires_provider else "0",
                )
            )
        )
        for provider in spec.provider_choices:
            print("\t".join(("PROVIDER", _safe(spec.treatment), _safe(provider))))

    patient_key = patient_id.strip().casefold()
    authoritative = {item.treatment.strip().casefold() for item in draft.treatments}
    for entry in read_base_schedule(workbook):
        if entry.patient_id.strip().casefold() != patient_key:
            continue
        if entry.treatment.strip().casefold() not in authoritative:
            continue
        print(
            "\t".join(
                (
                    "CURRENT",
                    _safe(entry.treatment),
                    _safe(entry.day_pattern),
                    entry.start_time.strftime("%H:%M"),
                    _safe(entry.therapist_id or ""),
                )
            )
        )

    for pattern in settings.day_patterns:
        print("\t".join(("DAY_PATTERN", _safe(pattern))))
    for slot in settings.standard_timeslots:
        print("\t".join(("STANDARD_SLOT", slot.strftime("%H:%M"))))
    for slot in settings.reclined_timeslots:
        print("\t".join(("RECLINED_SLOT", slot.strftime("%H:%M"))))


def _parse_request(raw: str) -> AutomaticTreatmentRequest:
    fields = raw.split("|", 3)
    while len(fields) < 4:
        fields.append("")
    treatment, day_pattern, raw_time, provider = (item.strip() for item in fields)
    if not treatment:
        raise ValueError("Automatic completion request is missing treatment")
    start_time = None
    if raw_time:
        start_time = datetime.strptime(raw_time, "%H:%M").time()
    return AutomaticTreatmentRequest(
        treatment=treatment,
        day_pattern=day_pattern or None,
        provider_id=provider or None,
        start_time=start_time,
    )


def _duration_minutes(treatment: str) -> int:
    if treatment.strip().casefold() == "Ανακλινόμενο".casefold():
        return RECLINED_TREATMENT_DURATION_MINUTES
    return STANDARD_TREATMENT_DURATION_MINUTES


def _time_minutes(value: time) -> int:
    return value.hour * 60 + value.minute


def _end_time_text(start_time: time, treatment: str) -> str:
    total = _time_minutes(start_time) + _duration_minutes(treatment)
    return f"{(total // 60) % 24:02d}:{total % 60:02d}"


def _overlap(left: AutomaticTreatmentRequest, right: AutomaticTreatmentRequest) -> bool:
    if left.start_time is None or right.start_time is None:
        return False
    left_start = _time_minutes(left.start_time)
    right_start = _time_minutes(right.start_time)
    left_end = left_start + _duration_minutes(left.treatment)
    right_end = right_start + _duration_minutes(right.treatment)
    return left_start < right_end and right_start < left_end


def _request_weekdays(request: AutomaticTreatmentRequest) -> frozenset[int]:
    if not request.day_pattern:
        return frozenset()
    return frozenset(int(day) for day in parse_day_pattern(request.day_pattern))


def _program_conflicts(
    requests: list[AutomaticTreatmentRequest],
) -> list[tuple[str, AutomaticTreatmentRequest, AutomaticTreatmentRequest]]:
    conflicts: list[tuple[str, AutomaticTreatmentRequest, AutomaticTreatmentRequest]] = []
    for left_index, left in enumerate(requests):
        if left.start_time is None or not left.day_pattern:
            continue
        left_days = _request_weekdays(left)
        for right in requests[left_index + 1 :]:
            if right.start_time is None or not right.day_pattern:
                continue
            shared_days = sorted(left_days.intersection(_request_weekdays(right)))
            if not shared_days or not _overlap(left, right):
                continue
            for weekday in shared_days:
                conflicts.append((_WEEKDAY_NAMES.get(weekday, str(weekday)), left, right))
    return conflicts


def _emit_validation(patient_id: str, raw_requests: list[str]) -> None:
    if not raw_requests:
        raise ValueError("At least one selected treatment is required")
    requests = [_parse_request(raw) for raw in raw_requests]
    print("\t".join(("OK", _safe(patient_id))))
    for weekday_name, left, right in _program_conflicts(requests):
        assert left.start_time is not None
        assert right.start_time is not None
        left_range = f"{left.start_time.strftime('%H:%M')}-{_end_time_text(left.start_time, left.treatment)}"
        right_range = f"{right.start_time.strftime('%H:%M')}-{_end_time_text(right.start_time, right.treatment)}"
        message = (
            f"{weekday_name}: {left.treatment} {left_range} συγκρούεται με "
            f"{right.treatment} {right_range}."
        )
        print(
            "\t".join(
                (
                    "CONFLICT",
                    _safe(weekday_name),
                    _safe(left.treatment),
                    left.start_time.strftime("%H:%M"),
                    _safe(right.treatment),
                    right.start_time.strftime("%H:%M"),
                    _safe(message),
                )
            )
        )


def _emit_completion(workbook: Path, patient_id: str, raw_requests: list[str]) -> None:
    if not raw_requests:
        raise ValueError("At least one selected treatment is required")

    requests = [_parse_request(raw) for raw in raw_requests]
    proposal = propose_automatic_patient_schedule(
        workbook,
        patient_id=patient_id,
        requests=requests,
        anchor_date=date.today(),
    )

    print("\t".join(("OK", _safe(patient_id))))
    for item in proposal.assignments:
        print(
            "\t".join(
                (
                    "ASSIGNMENT",
                    _safe(item.treatment),
                    _safe(item.day_pattern),
                    item.start_time.strftime("%H:%M"),
                    _safe(item.therapist_id),
                    _safe(item.therapist_name),
                    str(item.workload_count),
                )
            )
        )
    for item in proposal.unresolved:
        print(
            "\t".join(
                (
                    "UNRESOLVED",
                    _safe(item.treatment),
                    _safe(item.day_pattern),
                    _safe(item.reason),
                )
            )
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Patient program editor runtime bridge")
    parser.add_argument("command", choices=("config-vba", "complete-vba", "validate-vba"))
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--patient-id", required=True)
    parser.add_argument("--request", action="append", default=[])
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "config-vba":
            _emit_config(args.workbook, args.patient_id)
        elif args.command == "complete-vba":
            _emit_completion(args.workbook, args.patient_id, args.request)
        elif args.command == "validate-vba":
            _emit_validation(args.patient_id, args.request)
        return 0
    except Exception as exc:
        print("ERROR\t" + _safe(exc))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
