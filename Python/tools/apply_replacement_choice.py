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
from rehab_core.models import (  # noqa: E402
    ReplacementAssignment,
    ReplacementProviderKind,
    Therapist,
)
from rehab_core.replacement_options import find_replacement_options  # noqa: E402
from rehab_core.replacement_workflow import choose_replacement_option  # noqa: E402
from rehab_excel.native_excel import apply_write_plan_to_copy  # noqa: E402
from rehab_excel.patient_centric_preview import (  # noqa: E402
    PatientCentricPreviewError,
    build_patient_centric_preview_plan,
)
from rehab_excel.reader import (  # noqa: E402
    read_base_schedule,
    read_patients,
    read_settings,
)


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


def _is_legacy_student_name(name: str) -> bool:
    norm = _norm(name).replace(" ", "")
    return norm.startswith("ΦΟΙΤ") or norm.startswith("FOIT")


def _find_target_session(sessions, patients, patient_name: str, target_time: time):
    wanted = _norm(patient_name)
    patient_matches = [p for p in patients if _norm(p.display_name) == wanted]
    if len(patient_matches) != 1:
        raise ValueError(
            f"Expected one patient named {patient_name!r}; found {len(patient_matches)}"
        )
    patient = patient_matches[0]
    matches = [
        s
        for s in sessions
        if s.patient_id == patient.patient_id and s.start_time == target_time
    ]
    if len(matches) != 1:
        raise ValueError(
            f"Expected one active session for {patient.display_name} at "
            f"{target_time.strftime('%H:%M')}; found {len(matches)}"
        )
    return patient, matches[0]


def _fmt_times(slots: tuple[time, ...]) -> str:
    return ", ".join(slot.strftime("%H:%M") for slot in slots)


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def _ask_rank(option_count: int) -> int | None:
    while True:
        raw = input(f"Choose candidate [1-{option_count}] or 0 to cancel: ").strip()
        try:
            rank = int(raw)
        except ValueError:
            print("Enter a number.")
            continue
        if rank == 0:
            return None
        if 1 <= rank <= option_count:
            return rank
        print(f"Choose a number from 1 to {option_count}, or 0 to cancel.")


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Rank replacement providers, select one, then build a patient-centric "
            "Excel preview using that provider's recommended feasible timeslot."
        )
    )
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument("--patient", required=True)
    parser.add_argument("--time", type=_parse_time, required=True)
    parser.add_argument(
        "--rank",
        type=int,
        help="Use a ranked candidate directly. If omitted, the tool asks interactively.",
    )
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "GUIDED_REPLACEMENT_PREVIEW.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    try:
        patients = read_patients(source)
        settings = read_settings(source)
        base_entries = read_base_schedule(source)
        sessions = materialize_sessions_for_date(base_entries, args.date)
        patient, target = _find_target_session(
            sessions, patients, args.patient, args.time
        )
        if target.robotic:
            print(
                "SAFETY STOP: robotic capability is not yet authoritative in "
                "the workbook adapter; do not apply robotic replacements from this tool."
            )
            return 2

        therapists = [
            Therapist(therapist_id=name, display_name=name)
            for name in settings.therapist_names
            if not _is_legacy_student_name(name)
        ]
        options = find_replacement_options(
            target_session=target,
            therapists=therapists,
            sessions=sessions,
            patients=patients,
            requested_time=args.time,
            timeslots=settings.standard_timeslots,
        )
        if not options:
            print("SAFETY STOP: no feasible replacement providers/timeslots found")
            return 2

        visible = options[: max(args.limit, 1)]
        print("REPLACEMENT CANDIDATES")
        print(f"Date: {args.date.isoformat()}")
        print(f"Patient: {patient.display_name}")
        print(f"Original: {target.therapist_id} {target.start_time.strftime('%H:%M')}")
        for index, option in enumerate(visible, start=1):
            exact = "exact" if option.exact_time_available else "alternative"
            print(
                f"  {index}. {option.display_name} | load {option.active_sessions} | "
                f"recommended {option.recommended_time.strftime('%H:%M')} ({exact}) | "
                f"free: {_fmt_times(option.available_timeslots)}"
            )

        selected_rank = args.rank
        if selected_rank is None:
            selected_rank = _ask_rank(len(visible))
            if selected_rank is None:
                print("CANCELLED: no workbook was written.")
                return 0
        choice = choose_replacement_option(visible, selected_rank)
        option = choice.option

        replacement = ReplacementAssignment(
            replacement_id="guided-preview-1",
            target_session_id=target.session_id,
            patient_id=patient.patient_id,
            original_therapist_id=target.therapist_id,
            replacement_therapist_id=option.provider_id,
            replacement_date=args.date,
            replacement_time=option.recommended_time,
            reason="guided replacement preview",
            replacement_provider_kind=ReplacementProviderKind.THERAPIST,
        )
        states = build_daily_session_states(
            sessions,
            absences=(),
            replacements=(replacement,),
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
    print("GUIDED REPLACEMENT PREVIEW OK")
    print(
        f"Selected: #{choice.rank} {option.display_name} at "
        f"{option.recommended_time.strftime('%H:%M')}"
    )
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print("NEXT: open only the guided preview and inspect the original and destination cells.")
    return 0 if report.source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
