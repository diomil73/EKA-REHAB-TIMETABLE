from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.automatic_patient_schedule import (
    AutomaticTreatmentRequest,
    propose_automatic_patient_schedule,
)


def _parse_date(value: str) -> date:
    return date.fromisoformat(value)


def _json_dump(payload: dict) -> None:
    print(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))


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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Patient scheduling runtime bridge")
    parser.add_argument(
        "command",
        choices=("auto-propose", "auto-propose-vba"),
    )
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--patient-id", required=True)
    parser.add_argument("--date", required=True)
    parser.add_argument("--request", action="append", default=[])
    parser.add_argument("--robotic-capable", action="append", default=[])
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        payload = _auto_payload(args)
    except Exception as exc:
        if args.command.endswith("-vba"):
            print("ERROR\t" + _vba_safe(exc))
        else:
            _json_dump({"ok": False, "error": str(exc)})
        return 2

    if args.command == "auto-propose-vba":
        for line in _auto_vba_lines(payload):
            print(line)
    else:
        _json_dump(payload)
    return 0


if __name__ == "__main__":
    sys.exit(main())
