from __future__ import annotations

import argparse
import json
from datetime import date, datetime, time
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.automatic_patient_schedule import (
    AutomaticTreatmentRequest,
    propose_automatic_patient_schedule,
)
from rehab_excel.patient_recurring_scheduling import recurring_therapist_choices
from rehab_excel.patient_scheduling_flow import (
    TreatmentAssignment,
    refresh_and_validate_assignment,
    therapist_choices,
)
from rehab_excel.patient_treatment_plan import build_patient_treatment_plan


def _parse_date(value: str) -> date:
    return date.fromisoformat(value)


def _parse_time(value: str) -> time:
    return datetime.strptime(value, "%H:%M").time()


def _json_dump(payload: dict) -> None:
    print(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))


def _require_treatments(args: argparse.Namespace) -> None:
    if not args.treatment:
        raise ValueError("at least one --treatment is required")


def _plan_payload(args: argparse.Namespace) -> dict:
    _require_treatments(args)
    plan = build_patient_treatment_plan(
        args.workbook,
        patient_id=args.patient_id,
        selected_treatments=args.treatment,
        target_date=_parse_date(args.date),
        additional_robotic_capable_names=args.robotic_capable,
    )
    treatments = []
    for item in plan.treatments:
        choices = therapist_choices(plan, item.treatment)
        treatments.append(
            {
                "treatment": item.treatment,
                "requires_therapist": item.requires_therapist,
                "therapists": [
                    {
                        "therapist_id": choice.therapist_id,
                        "label": choice.label,
                        "available_slots": [slot.strftime("%H:%M") for slot in choice.available_slots],
                    }
                    for choice in choices
                ],
            }
        )
    return {
        "ok": True,
        "patient_id": plan.patient_id,
        "date": plan.target_date.isoformat(),
        "treatments": treatments,
    }


def _validate_payload(args: argparse.Namespace) -> dict:
    _require_treatments(args)
    result = refresh_and_validate_assignment(
        args.workbook,
        patient_id=args.patient_id,
        selected_treatments=args.treatment,
        target_date=_parse_date(args.date),
        assignment=TreatmentAssignment(
            treatment=args.assignment_treatment,
            therapist_id=args.therapist_id,
            start_time=_parse_time(args.start_time),
        ),
        additional_robotic_capable_names=args.robotic_capable,
    )
    return {
        "ok": True,
        "valid": result.valid,
        "message": result.message,
    }


def _recurring_payload(args: argparse.Namespace) -> dict:
    _require_treatments(args)
    if not args.day_pattern:
        raise ValueError("recurring scheduling requires --day-pattern")
    treatment = args.treatment[-1]
    choices = recurring_therapist_choices(
        args.workbook,
        patient_id=args.patient_id,
        treatment=treatment,
        day_pattern=args.day_pattern,
        anchor_date=_parse_date(args.date),
        additional_robotic_capable_names=args.robotic_capable,
    )
    return {
        "ok": True,
        "patient_id": args.patient_id,
        "date": args.date,
        "treatment": treatment,
        "day_pattern": args.day_pattern,
        "therapists": [
            {
                "therapist_id": choice.therapist_id,
                "label": choice.workload_label,
                "available_slots": [slot.strftime("%H:%M") for slot in choice.available_slots],
            }
            for choice in choices
        ],
    }


def _recurring_validate_payload(args: argparse.Namespace) -> dict:
    payload = _recurring_payload(args)
    if not args.therapist_id or not args.start_time:
        raise ValueError("validate-recurring-vba requires --therapist-id and --start-time")
    selected_time = _parse_time(args.start_time).strftime("%H:%M")
    for therapist in payload["therapists"]:
        if therapist["therapist_id"] != args.therapist_id:
            continue
        valid = selected_time in therapist["available_slots"]
        return {
            "ok": True,
            "valid": valid,
            "message": "OK" if valid else "Η επιλεγμένη ώρα δεν είναι πλέον διαθέσιμη.",
        }
    return {
        "ok": True,
        "valid": False,
        "message": "Ο επιλεγμένος θεραπευτής δεν είναι πλέον διαθέσιμος για το μοτίβο ημερών.",
    }


def _parse_auto_request(value: str) -> AutomaticTreatmentRequest:
    treatment, separator, day_pattern = value.partition("|")
    if not separator or not treatment.strip() or not day_pattern.strip():
        raise ValueError("--request must be TREATMENT|DAY_PATTERN")
    return AutomaticTreatmentRequest(treatment.strip(), day_pattern.strip())


def _auto_payload(args: argparse.Namespace) -> dict:
    if not args.request:
        raise ValueError("auto-propose requires at least one --request TREATMENT|DAY_PATTERN")
    proposal = propose_automatic_patient_schedule(
        args.workbook,
        patient_id=args.patient_id,
        requests=tuple(_parse_auto_request(value) for value in args.request),
        anchor_date=_parse_date(args.date),
        additional_robotic_capable_names=args.robotic_capable,
    )
    return {
        "ok": True,
        "patient_id": proposal.patient_id,
        "date": proposal.anchor_date.isoformat(),
        "complete": proposal.complete,
        "total_workload": proposal.total_workload,
        "assignments": [
            {
                "treatment": item.treatment,
                "day_pattern": item.day_pattern,
                "therapist_id": item.therapist_id,
                "therapist_name": item.therapist_name,
                "start_time": item.start_time.strftime("%H:%M"),
                "workload_count": item.workload_count,
            }
            for item in proposal.assignments
        ],
        "unresolved": [
            {
                "treatment": item.treatment,
                "day_pattern": item.day_pattern,
                "reason": item.reason,
            }
            for item in proposal.unresolved
        ],
    }


def _vba_safe(value: object) -> str:
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def _plan_vba_lines(payload: dict) -> tuple[str, ...]:
    lines = ["\t".join(("OK", _vba_safe(payload["patient_id"]), _vba_safe(payload["date"])))]
    for treatment in payload["treatments"]:
        lines.append(
            "\t".join(
                (
                    "TREATMENT",
                    _vba_safe(treatment["treatment"]),
                    "1" if treatment["requires_therapist"] else "0",
                )
            )
        )
        for therapist in treatment["therapists"]:
            lines.append(
                "\t".join(
                    (
                        "THERAPIST",
                        _vba_safe(treatment["treatment"]),
                        _vba_safe(therapist["therapist_id"]),
                        _vba_safe(therapist["label"]),
                        "|".join(_vba_safe(slot) for slot in therapist["available_slots"]),
                    )
                )
            )
    return tuple(lines)


def _recurring_vba_lines(payload: dict) -> tuple[str, ...]:
    lines = [
        "\t".join(
            (
                "OK",
                _vba_safe(payload["patient_id"]),
                _vba_safe(payload["date"]),
                _vba_safe(payload["treatment"]),
                _vba_safe(payload["day_pattern"]),
            )
        )
    ]
    for therapist in payload["therapists"]:
        lines.append(
            "\t".join(
                (
                    "THERAPIST",
                    _vba_safe(payload["treatment"]),
                    _vba_safe(therapist["therapist_id"]),
                    _vba_safe(therapist["label"]),
                    "|".join(_vba_safe(slot) for slot in therapist["available_slots"]),
                )
            )
        )
    return tuple(lines)


def _auto_vba_lines(payload: dict) -> tuple[str, ...]:
    lines = [
        "\t".join(
            (
                "OK",
                _vba_safe(payload["patient_id"]),
                _vba_safe(payload["date"]),
                "1" if payload["complete"] else "0",
                _vba_safe(payload["total_workload"]),
            )
        )
    ]
    for item in payload["assignments"]:
        lines.append(
            "\t".join(
                (
                    "ASSIGNMENT",
                    _vba_safe(item["treatment"]),
                    _vba_safe(item["day_pattern"]),
                    _vba_safe(item["therapist_id"]),
                    _vba_safe(item["therapist_name"]),
                    _vba_safe(item["start_time"]),
                    _vba_safe(item["workload_count"]),
                )
            )
        )
    for item in payload["unresolved"]:
        lines.append(
            "\t".join(
                (
                    "UNRESOLVED",
                    _vba_safe(item["treatment"]),
                    _vba_safe(item["day_pattern"]),
                    _vba_safe(item["reason"]),
                )
            )
        )
    return tuple(lines)


def _validate_vba_line(payload: dict) -> str:
    return "\t".join(("OK", "1" if payload["valid"] else "0", _vba_safe(payload["message"])))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Patient scheduling runtime bridge")
    parser.add_argument(
        "command",
        choices=(
            "plan",
            "validate",
            "plan-vba",
            "validate-vba",
            "recurring-vba",
            "validate-recurring-vba",
            "auto-propose",
            "auto-propose-vba",
        ),
    )
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--patient-id", required=True)
    parser.add_argument("--date", required=True)
    parser.add_argument("--treatment", action="append", default=[])
    parser.add_argument("--request", action="append", default=[])
    parser.add_argument("--day-pattern")
    parser.add_argument("--robotic-capable", action="append", default=[])
    parser.add_argument("--assignment-treatment")
    parser.add_argument("--therapist-id")
    parser.add_argument("--start-time")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command in {"auto-propose", "auto-propose-vba"}:
            payload = _auto_payload(args)
        elif args.command in {"plan", "plan-vba"}:
            payload = _plan_payload(args)
        elif args.command in {"recurring-vba", "validate-recurring-vba"}:
            payload = (
                _recurring_payload(args)
                if args.command == "recurring-vba"
                else _recurring_validate_payload(args)
            )
        else:
            missing = [
                name
                for name, value in (
                    ("--assignment-treatment", args.assignment_treatment),
                    ("--therapist-id", args.therapist_id),
                    ("--start-time", args.start_time),
                )
                if not value
            ]
            if missing:
                raise ValueError("validate requires " + ", ".join(missing))
            payload = _validate_payload(args)
    except Exception as exc:
        if args.command.endswith("-vba"):
            print("ERROR\t" + _vba_safe(exc))
        else:
            _json_dump({"ok": False, "error": str(exc)})
        return 2

    if args.command == "plan-vba":
        for line in _plan_vba_lines(payload):
            print(line)
    elif args.command == "recurring-vba":
        for line in _recurring_vba_lines(payload):
            print(line)
    elif args.command == "auto-propose-vba":
        for line in _auto_vba_lines(payload):
            print(line)
    elif args.command in {"validate-vba", "validate-recurring-vba"}:
        print(_validate_vba_line(payload))
    else:
        _json_dump(payload)
    return 0


if __name__ == "__main__":
    sys.exit(main())
