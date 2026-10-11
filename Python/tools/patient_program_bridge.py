from __future__ import annotations

import argparse
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

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
            print(
                "\t".join(
                    (
                        "PROVIDER",
                        _safe(spec.treatment),
                        _safe(provider),
                    )
                )
            )

    for pattern in settings.day_patterns:
        print("\t".join(("DAY_PATTERN", _safe(pattern))))
    for slot in settings.standard_timeslots:
        print("\t".join(("STANDARD_SLOT", slot.strftime("%H:%M"))))
    for slot in settings.reclined_timeslots:
        print("\t".join(("RECLINED_SLOT", slot.strftime("%H:%M"))))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Patient program editor runtime bridge")
    parser.add_argument("command", choices=("config-vba",))
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--patient-id", required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "config-vba":
            _emit_config(args.workbook, args.patient_id)
        return 0
    except Exception as exc:
        print("ERROR\t" + _safe(exc))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
