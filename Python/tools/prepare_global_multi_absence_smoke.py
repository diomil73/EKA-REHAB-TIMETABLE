from __future__ import annotations

import argparse
import sys
from collections import Counter
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_excel.application_contract import inspect_application_contract  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_patients import audit_daily_input_patients  # noqa: E402
from rehab_excel.daily_input_sheet import (  # noqa: E402
    DailyInputSheetError,
    build_daily_input_spec,
    create_daily_input_preview,
)
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402


class MultiAbsenceSmokeError(RuntimeError):
    pass


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def choose_smoke_therapists(
    sessions,
    *,
    allowed_therapists: tuple[str, ...],
    therapist_count: int = 2,
    min_sessions_per_therapist: int = 1,
) -> tuple[tuple[str, tuple], ...]:
    """Choose multiple real providers while keeping the smoke load modest.

    Prefer therapists with the fewest qualifying sessions so the scenario is a true
    multi-absence case without needlessly making a fully-resolved smoke impossible.
    """
    if therapist_count < 2:
        raise ValueError("therapist_count must be at least 2")
    if min_sessions_per_therapist < 1:
        raise ValueError("min_sessions_per_therapist must be positive")

    allowed = set(allowed_therapists)
    eligible = tuple(session for session in sessions if session.therapist_id in allowed)
    counts = Counter(session.therapist_id for session in eligible)
    candidates = sorted(
        (
            (count, therapist_id)
            for therapist_id, count in counts.items()
            if count >= min_sessions_per_therapist
        ),
        key=lambda item: (item[0], item[1]),
    )
    if len(candidates) < therapist_count:
        raise ValueError(
            f"Need {therapist_count} therapists with at least "
            f"{min_sessions_per_therapist} session(s); found {len(candidates)}"
        )

    selected: list[tuple[str, tuple]] = []
    for _count, therapist_id in candidates[:therapist_count]:
        affected = tuple(
            sorted(
                (session for session in eligible if session.therapist_id == therapist_id),
                key=lambda session: (session.start_time, session.patient_id, session.session_id),
            )
        )
        selected.append((therapist_id, affected))
    return tuple(selected)


def _write_absence_rows(workbook_path: Path, therapist_ids: tuple[str, ...]) -> None:
    if sys.platform != "win32":
        raise MultiAbsenceSmokeError("Windows + Microsoft Excel are required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise MultiAbsenceSmokeError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(
            str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False
        )
        ws = workbook.Worksheets("DAILY_INPUT")
        for offset, therapist_id in enumerate(therapist_ids):
            row = 7 + offset
            ws.Cells(row, 1).Value = therapist_id
            ws.Cells(row, 2).Value = "ΝΑΙ"
            ws.Range(ws.Cells(row, 3), ws.Cells(row, 4)).ClearContents()
            ws.Cells(row, 5).Value = "Global multi-absence smoke"
            ws.Cells(row, 6).Value = "Resolve all displaced sessions together"
        workbook.Save()
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a full-app DAILY_INPUT smoke copy with multiple therapist absences."
    )
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "OUTPATIENT_RECURRING_SMOKE.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "GLOBAL_MULTI_ABSENCE_INPUT.xlsm",
    )
    parser.add_argument("--therapists", type=int, default=2)
    parser.add_argument("--min-sessions", type=int, default=1)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    source_hash = file_sha256(source)
    try:
        source_contract = inspect_application_contract(source)
        if not source_contract.ok:
            raise MultiAbsenceSmokeError(
                "Source is not a complete application workbook: "
                + ", ".join(source_contract.missing_components)
            )

        patient_audit = audit_daily_input_patients(source)
        if not patient_audit.ok:
            raise MultiAbsenceSmokeError("Patient identity audit failed")
        settings = read_settings(source)
        base_entries = read_unified_base_schedule(source)
        sessions = materialize_sessions_for_date(base_entries, args.date)
        selected = choose_smoke_therapists(
            sessions,
            allowed_therapists=tuple(settings.therapist_names),
            therapist_count=args.therapists,
            min_sessions_per_therapist=args.min_sessions,
        )
        patients = read_patient_registry(source)
        patient_by_id = {patient.patient_id: patient.display_name for patient in patients}

        spec = build_daily_input_spec(
            target_date=args.date,
            therapist_names=settings.therapist_names,
            patient_names=patient_audit.dropdown_names,
            patient_statuses=settings.patient_statuses,
            timeslots=settings.standard_timeslots,
        )
        preview, _ = create_daily_input_preview(
            source, output, spec, overwrite=args.overwrite
        )
        _write_absence_rows(preview, tuple(item[0] for item in selected))

        output_contract = inspect_application_contract(output)
        if not output_contract.ok:
            raise MultiAbsenceSmokeError(
                "Output lost application components: "
                + ", ".join(output_contract.missing_components)
            )
    except (ValueError, DailyInputSheetError, MultiAbsenceSmokeError) as exc:
        try:
            output.unlink(missing_ok=True)
        except OSError:
            pass
        print(f"SAFETY STOP: {exc}")
        return 2

    source_unchanged = file_sha256(source) == source_hash
    print("GLOBAL MULTI-ABSENCE SMOKE INPUT OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Absent therapists: {len(selected)}")
    total = 0
    for therapist_id, affected in selected:
        print(f"  {therapist_id}: {len(affected)} affected session(s)")
        total += len(affected)
        for session in affected:
            patient_name = patient_by_id.get(session.patient_id, session.patient_id)
            print(f"    {session.start_time.strftime('%H:%M')} | {patient_name}")
    print(f"Total displaced sessions: {total}")
    print(f"Preview: {output}")
    print(f"Source unchanged: {source_unchanged}")
    print("Application contract output: OK")
    return 0 if source_unchanged else 3


if __name__ == "__main__":
    raise SystemExit(main())
