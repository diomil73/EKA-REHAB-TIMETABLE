from __future__ import annotations

import argparse
import sys
import unicodedata
from datetime import date, datetime, time, timedelta
from pathlib import Path
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_core.models import (  # noqa: E402
    AbsenceKind,
    DailyAbsence,
    ReplacementAssignment,
    ReplacementProviderKind,
)
from rehab_excel.native_excel import apply_write_plan_to_copy  # noqa: E402
from rehab_excel.operational_overlay import (  # noqa: E402
    OperationalOverlayError,
    build_operational_overlay_write_plan,
)
from rehab_excel.reader import read_base_schedule, read_patients  # noqa: E402


def _norm(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().upper())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid time {value!r}; use HH:MM") from exc


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid date {value!r}; use YYYY-MM-DD") from exc


def _one_minute_after(value: time) -> time:
    stamp = datetime.combine(date(2000, 1, 1), value) + timedelta(minutes=1)
    return stamp.time()


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
        candidates = ", ".join(
            f"{session.therapist_id}@{session.start_time.strftime('%H:%M')}"
            for session in sessions
            if session.patient_id == patient_id
        ) or "none"
        raise ValueError(
            f"Expected one session for patient id {patient_id!r} at "
            f"{original_time.strftime('%H:%M')}; found {len(matches)}. Candidates: {candidates}"
        )
    return matches[0]


def _provider_id(sessions, provider_name: str) -> str:
    wanted = _norm(provider_name)
    known = {session.therapist_id for session in sessions}
    matches = [provider for provider in known if _norm(provider) == wanted]
    if len(matches) == 1:
        return matches[0]
    # A replacement provider can have no base session that day. In that case
    # keep the user's exact label and let the verified THERAPIST_DAILY mapper
    # decide whether that provider column exists.
    return provider_name.strip()


def _parse_replacement(spec: str):
    parts = [item.strip() for item in spec.split("|")]
    if len(parts) != 4:
        raise ValueError(
            "Replacement must be PATIENT|ORIGINAL_TIME|REPLACEMENT_PROVIDER|REPLACEMENT_TIME"
        )
    return parts[0], _parse_time(parts[1]), parts[2], _parse_time(parts[3])


def _parse_patient_absence(spec: str):
    parts = [item.strip() for item in spec.split("|", 2)]
    if len(parts) not in {2, 3}:
        raise ValueError("Patient absence must be PATIENT|TIME_OR_ALL[|REASON]")
    return parts[0], parts[1], parts[2] if len(parts) == 3 else None


def _parse_therapist_absence(spec: str):
    parts = [item.strip() for item in spec.split("|", 2)]
    if len(parts) not in {2, 3}:
        raise ValueError("Therapist absence must be PROVIDER|ALL_OR_START-END[|REASON]")
    return parts[0], parts[1], parts[2] if len(parts) == 3 else None


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def build_inputs(source: Path, target_date: date, replacement_specs, patient_absence_specs, therapist_absence_specs):
    patients = read_patients(source)
    base_entries = read_base_schedule(source)
    sessions = materialize_sessions_for_date(base_entries, target_date)

    replacements: list[ReplacementAssignment] = []
    absences: list[DailyAbsence] = []

    for index, raw in enumerate(replacement_specs, start=1):
        patient_name, original_time, provider_name, replacement_time = _parse_replacement(raw)
        patient = _unique_patient(patients, patient_name)
        session = _unique_session(sessions, patient.patient_id, original_time)
        replacement_provider = _provider_id(sessions, provider_name)
        replacements.append(
            ReplacementAssignment(
                replacement_id=f"preview-replacement-{index}",
                target_session_id=session.session_id,
                patient_id=patient.patient_id,
                original_therapist_id=session.therapist_id,
                replacement_therapist_id=replacement_provider,
                replacement_date=target_date,
                replacement_time=replacement_time,
                reason="operational preview",
                replacement_provider_kind=ReplacementProviderKind.THERAPIST,
            )
        )

    for raw in patient_absence_specs:
        patient_name, slot, reason = _parse_patient_absence(raw)
        patient = _unique_patient(patients, patient_name)
        if _norm(slot) == "ALL":
            start = end = None
        else:
            start = _parse_time(slot)
            end = _one_minute_after(start)
        absences.append(
            DailyAbsence(
                absence_kind=AbsenceKind.PATIENT,
                subject_id=patient.patient_id,
                absence_date=target_date,
                start_time=start,
                end_time=end,
                reason=reason,
            )
        )

    for raw in therapist_absence_specs:
        provider_name, period, reason = _parse_therapist_absence(raw)
        provider_id = _provider_id(sessions, provider_name)
        if _norm(period) == "ALL":
            start = end = None
        else:
            if "-" not in period:
                raise ValueError("Therapist period must be ALL or HH:MM-HH:MM")
            start_raw, end_raw = period.split("-", 1)
            start, end = _parse_time(start_raw), _parse_time(end_raw)
        absences.append(
            DailyAbsence(
                absence_kind=AbsenceKind.THERAPIST,
                subject_id=provider_id,
                absence_date=target_date,
                start_time=start,
                end_time=end,
                reason=reason,
            )
        )

    return patients, sessions, absences, replacements


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create a generic daily operational overlay preview on a NEW .xlsm copy. "
            "No patient data is exported to sidecar files."
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
        "--patient-absence",
        action="append",
        default=[],
        metavar="PATIENT|TIME_OR_ALL|REASON",
    )
    parser.add_argument(
        "--therapist-absence",
        action="append",
        default=[],
        metavar="PROVIDER|ALL_OR_START-END|REASON",
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=(
            REPO_ROOT
            / "Excel"
            / "previews"
            / "Rehab_Center_System_v27_1_DAILY_OVERLAY_PREVIEW.xlsm"
        ),
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    if not (args.replace or args.patient_absence or args.therapist_absence):
        print("SAFETY STOP: provide at least one daily operation")
        return 2

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    try:
        patients, sessions, absences, replacements = build_inputs(
            source,
            args.date,
            args.replace,
            args.patient_absence,
            args.therapist_absence,
        )
        states = build_daily_session_states(
            sessions,
            absences,
            replacements,
            target_date=args.date,
        )
        overlay = build_operational_overlay_write_plan(
            source,
            states,
            sessions,
            patients,
        )
        report = apply_write_plan_to_copy(
            overlay.write_plan,
            output,
            overwrite=args.overwrite,
        )
    except (ValueError, OperationalOverlayError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    vba_preserved = _has_vba(source) == _has_vba(output) and _has_vba(output)
    print("DAILY OVERLAY PREVIEW OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Changed sessions: {len(overlay.bindings)}")
    print(f"Touched cells: {report.operation_count}")
    for binding in overlay.bindings:
        suffix = f" -> {binding.effective_cell}" if binding.effective_cell else ""
        print(
            f"  {binding.status.value}: {binding.session_id} "
            f"[{binding.original_cell}{suffix}]"
        )
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print("NEXT: open ONLY the daily overlay preview and inspect the touched cells.")
    return 0 if report.source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
