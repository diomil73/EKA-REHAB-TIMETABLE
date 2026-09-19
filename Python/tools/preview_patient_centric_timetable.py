from __future__ import annotations

import argparse
import sys
import unicodedata
from datetime import date, datetime, time
from pathlib import Path
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_core.models import ReplacementAssignment, ReplacementProviderKind  # noqa: E402
from rehab_excel.native_excel import apply_write_plan_to_copy  # noqa: E402
from rehab_excel.patient_centric_preview import (  # noqa: E402
    PatientCentricPreviewError,
    build_patient_centric_preview_plan,
)
from rehab_excel.reader import read_base_schedule, read_patients  # noqa: E402


def _norm(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().upper())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Time must be HH:MM") from exc


def _parse_replacement(value: str):
    parts = [part.strip() for part in value.split("|")]
    if len(parts) != 4:
        raise ValueError("Replacement must be PATIENT|ORIGINAL_TIME|PROVIDER|NEW_TIME")
    return parts[0], _parse_time(parts[1]), parts[2], _parse_time(parts[3])


def _unique_patient(patients, name: str):
    wanted = _norm(name)
    matches = [patient for patient in patients if _norm(patient.display_name) == wanted]
    if len(matches) != 1:
        raise ValueError(f"Expected one patient named {name!r}; found {len(matches)}")
    return matches[0]


def _unique_session(sessions, patient_id: str, original_time: time):
    matches = [
        session
        for session in sessions
        if session.patient_id == patient_id and session.start_time == original_time
    ]
    if len(matches) != 1:
        raise ValueError(
            f"Expected one dated session for patient {patient_id!r} at "
            f"{original_time.strftime('%H:%M')}; found {len(matches)}"
        )
    return matches[0]


def _provider_id(sessions, name: str) -> str:
    wanted = _norm(name)
    providers = {session.therapist_id for session in sessions}
    matches = [provider for provider in providers if _norm(provider) == wanted]
    return matches[0] if len(matches) == 1 else name.strip()


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Build a patient-centric THERAPIST_DAILY preview. Pair/group cells are "
            "rebuilt from independent patient assignments; same-patient split schedules "
            "show the patient name once."
        )
    )
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument(
        "--replace",
        action="append",
        default=[],
        metavar="PATIENT|ORIGINAL_TIME|PROVIDER|NEW_TIME",
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "PATIENT_CENTRIC_TIMETABLE_PREVIEW.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    if not args.replace:
        print("SAFETY STOP: provide at least one --replace operation for this preview")
        return 2

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    try:
        patients = read_patients(source)
        base_entries = read_base_schedule(source)
        sessions = materialize_sessions_for_date(base_entries, args.date)
        replacements: list[ReplacementAssignment] = []
        for index, raw in enumerate(args.replace, start=1):
            patient_name, original_time, provider_name, new_time = _parse_replacement(raw)
            patient = _unique_patient(patients, patient_name)
            session = _unique_session(sessions, patient.patient_id, original_time)
            provider_id = _provider_id(sessions, provider_name)
            replacements.append(
                ReplacementAssignment(
                    replacement_id=f"patient-centric-preview-{index}",
                    target_session_id=session.session_id,
                    patient_id=patient.patient_id,
                    original_therapist_id=session.therapist_id,
                    replacement_therapist_id=provider_id,
                    replacement_date=args.date,
                    replacement_time=new_time,
                    reason="patient-centric visual preview",
                    replacement_provider_kind=ReplacementProviderKind.THERAPIST,
                )
            )

        states = build_daily_session_states(
            sessions,
            absences=(),
            replacements=replacements,
            target_date=args.date,
        )
        preview = build_patient_centric_preview_plan(
            source,
            base_entries=base_entries,
            states=states,
            sessions=sessions,
            patients=patients,
            rebuild_multi_member_groups=True,
        )
        report = apply_write_plan_to_copy(
            preview.write_plan,
            output,
            overwrite=args.overwrite,
        )
    except (ValueError, PatientCentricPreviewError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    vba_preserved = _has_vba(source) == _has_vba(output) and _has_vba(output)
    print("PATIENT-CENTRIC TIMETABLE PREVIEW OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Rebuilt multi-assignment cells: {len(preview.rebuilt_group_cells)}")
    if preview.rebuilt_group_cells:
        print("  " + ", ".join(preview.rebuilt_group_cells))
    for binding in preview.bindings:
        suffix = f" -> {binding.effective_cell}" if binding.effective_cell else ""
        print(f"  {binding.status.value}: {binding.session_id} [{binding.original_cell}{suffix}]")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print("NEXT: inspect paired cells, same-patient split cells, and the replacement destination.")
    return 0 if report.source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
