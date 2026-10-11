from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.automatic_patient_schedule import (
    AutomaticTreatmentRequest,
    propose_automatic_patient_schedule,
)
from rehab_excel.patient_program_draft import build_patient_program_draft
from rehab_excel.reader import read_settings


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
    parser.add_argument("command", choices=("config-vba", "complete-vba"))
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
        return 0
    except Exception as exc:
        print("ERROR\t" + _safe(exc))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
