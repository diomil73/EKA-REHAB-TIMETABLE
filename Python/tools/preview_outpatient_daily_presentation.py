from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_excel.native_excel import apply_write_plan_to_copy  # noqa: E402
from rehab_excel.outpatient_presentation import (  # noqa: E402
    OutpatientPresentationError,
    build_outpatient_daily_patches,
)
from rehab_excel.outpatient_schedule_source import (  # noqa: E402
    OutpatientScheduleSourceError,
    read_unified_base_schedule,
)
from rehab_excel.patient_registry_source import (  # noqa: E402
    PatientRegistrySourceError,
    read_patient_registry,
)
from rehab_excel.writeback import build_write_plan  # noqa: E402


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create a safe copy applying only active outpatient THERAPIST_DAILY "
            "presentation for one date."
        )
    )
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    try:
        patients = read_patient_registry(source)
        entries = read_unified_base_schedule(source)
        sessions = materialize_sessions_for_date(entries, args.date)
        states = build_daily_session_states(
            sessions,
            absences=(),
            replacements=(),
            cancellations=(),
            target_date=args.date,
        )
        patches = build_outpatient_daily_patches(
            source,
            states=states,
            patients=patients,
        )
        if not patches:
            raise OutpatientPresentationError(
                "No outpatient presentation patches were produced for the selected date"
            )
        plan = build_write_plan(source, patches)
        report = apply_write_plan_to_copy(plan, output, overwrite=args.overwrite)
    except (
        OutpatientPresentationError,
        OutpatientScheduleSourceError,
        PatientRegistrySourceError,
        ValueError,
        KeyError,
    ) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    vba_preserved = _has_vba(source) == _has_vba(output) and _has_vba(output)
    blue_targets = [
        patch for patch in patches if patch.fill_role == "outpatient_light_blue"
    ]
    clear_targets = [patch for patch in patches if patch.fill_role == "clear_fill"]

    print("OUTPATIENT DAILY PRESENTATION PREVIEW OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Blue outpatient cells: {len(blue_targets)}")
    for patch in blue_targets:
        print(f"  {patch.sheet}!{patch.cell} -> DDEBF7")
    print(f"Stale blue cells cleared: {len(clear_targets)}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print(f"NEXT: open only {output.name} and inspect the listed blue cell(s).")
    return 0 if report.source_unchanged and vba_preserved and blue_targets else 3


if __name__ == "__main__":
    raise SystemExit(main())
