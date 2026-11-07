from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.patient_history_reports import (  # noqa: E402
    read_patient_history_report,
    include_registered_patients_without_sessions,
)
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402


HEADERS = (
    "PeriodStart", "PeriodEnd", "PatientId", "PatientName", "FinalizedDays",
    "Scheduled", "Performed", "PatientAbsences", "Cancellations", "Replacements",
    "CompletionPct",
)


def _display_patient_id(value: str) -> str:
    """Pad one-digit numeric identifiers only for display; keep stored IDs intact."""
    return value.zfill(2) if value.isascii() and value.isdecimal() and len(value) == 1 else value


def _identity(patient_id: object) -> tuple[str, str | int]:
    """Match stable workbook IDs without conflating arbitrary text identifiers."""
    clean = str(patient_id).strip()
    if not clean:
        raise ValueError("Blank patient identifier")
    if clean.isascii() and clean.isdecimal():
        return ("numeric", int(clean))
    return ("text", clean.casefold())


def _patient_names(registered: list) -> dict[tuple[str, str | int], str]:
    """Only current verified registry names; never infer historical names."""
    names: dict[tuple[str, str | int], str] = {}
    for patient in registered:
        key = _identity(patient.patient_id)
        if key in names:
            raise ValueError(f"Duplicate patient identifier in registry: {patient.patient_id}")
        label = str(patient.display_name or "").strip()
        if not label:
            raise ValueError(f"Patient is missing display name: {patient.patient_id}")
        names[key] = label
    return names


def _completion_pct(performed: int, scheduled: int) -> str:
    """Delivered / all scheduled finalized sessions, including cancellations.

    This is a completion/participation-on-scheduled measure, not attendance
    against an adjusted eligible denominator. The cancellation-policy KPI is
    a separate future metric and must not be silently substituted here.
    """
    if scheduled < 0 or performed < 0 or performed > scheduled:
        raise ValueError("Invalid scheduled/performed values for completion percentage")
    if scheduled == 0:
        return ""
    return f"{performed * 100 / scheduled:.1f}"


def _field(value: object) -> str:
    text = str(value)
    if any(c in text for c in ("\t", "\r", "\n")):
        raise ValueError("Patient history report contains an unsupported control character")
    return text


def export_patient_history_ui(
    workbook_path: str | Path,
    output_path: str | Path,
    *,
    mode: str = "monthly",
    start: date | None = None,
    end: date | None = None,
    include_registered_zero: bool = False,
) -> dict[str, object]:
    """Emit report TSV without opening or changing Excel."""
    rows = read_patient_history_report(workbook_path, mode=mode, start=start, end=end)
    registered = read_patient_registry(workbook_path)
    patient_names = _patient_names(registered)
    if include_registered_zero:
        rows = include_registered_patients_without_sessions(
            rows, tuple(str(patient.patient_id) for patient in registered), mode=mode
        )
    lines = ["\t".join(HEADERS)]
    for row in rows:
        fields = (
            row.period_start.isoformat(), row.period_end.isoformat(),
            _display_patient_id(row.patient_id),
            patient_names.get(
                _identity(row.patient_id),
                f"Άγνωστο όνομα (κωδικός {_display_patient_id(row.patient_id)})",
            ),
            row.finalized_days, row.scheduled, row.performed,
            row.patient_absences, row.cancellations, row.replacements,
            _completion_pct(row.performed, row.scheduled),
        )
        lines.append("\t".join(_field(item) for item in fields))
    output = Path(output_path).resolve()
    if output.exists():
        raise ValueError("Output already exists; refusing overwrite")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return {"ok": True, "mode": mode, "rows": len(rows), "output": str(output)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only H7 patient history report UI TSV")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--mode", choices=("daily", "weekly", "monthly", "overall"), default="monthly")
    parser.add_argument("--start", type=date.fromisoformat)
    parser.add_argument("--end", type=date.fromisoformat)
    parser.add_argument("--include-registered-zero", action="store_true",
                        help="Show all current registry patients, including those without finalized activity")
    args = parser.parse_args()
    try:
        result = export_patient_history_ui(
            args.workbook, args.output, mode=args.mode, start=args.start, end=args.end,
            include_registered_zero=args.include_registered_zero
        )
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    print("PATIENT HISTORY UI EXPORT OK")
    print(f"Mode: {result['mode']}")
    print(f"Rows: {result['rows']}")
    print(f"Output: {result['output']}")
    print("Workbook unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
