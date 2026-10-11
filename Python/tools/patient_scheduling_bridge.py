from __future__ import annotations

import argparse
import json
from datetime import date, datetime, time
from pathlib import Path
import sys

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


def _plan_payload(args: argparse.Namespace) -> dict:
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


def _vba_safe(value: object) -> str:
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def _plan_vba_lines(payload: dict) -> tuple[str, ...]:
    lines = [
        "\t".join(
            (
                "OK",
                _vba_safe(payload["patient_id"]),
                _vba_safe(payload["date"]),
            )
        )
    ]
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


def _validate_vba_line(payload: dict) -> str:
    return "\t".join(
        (
            "OK",
            "1" if payload["valid"] else "0",
            _vba_safe(payload["message"]),
        )
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Patient scheduling runtime bridge")
    parser.add_argument(
        "command",
        choices=("plan", "validate", "plan-vba", "validate-vba"),
    )
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--patient-id", required=True)
    parser.add_argument("--date", required=True)
    parser.add_argument("--treatment", action="append", required=True)
    parser.add_argument("--robotic-capable", action="append", default=[])
    parser.add_argument("--assignment-treatment")
    parser.add_argument("--therapist-id")
    parser.add_argument("--start-time")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command in {"plan", "plan-vba"}:
            payload = _plan_payload(args)
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
    elif args.command == "validate-vba":
        print(_validate_vba_line(payload))
    else:
        _json_dump(payload)
    return 0


if __name__ == "__main__":
    sys.exit(main())
