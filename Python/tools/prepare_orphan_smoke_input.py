from __future__ import annotations

import argparse
import sys
from collections import Counter
from datetime import date
from pathlib import Path
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
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


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def choose_smoke_therapist(
    sessions,
    *,
    allowed_therapists: tuple[str, ...],
    min_sessions: int = 2,
) -> tuple[str, tuple]:
    """Pick a therapist whose absence creates a useful multi-orphan smoke case."""

    if min_sessions < 2:
        raise ValueError("min_sessions must be at least 2 for a mixed orphan smoke case")

    allowed = set(allowed_therapists)
    eligible = [session for session in sessions if session.therapist_id in allowed]
    counts = Counter(session.therapist_id for session in eligible)
    candidates = sorted(
        (
            (count, therapist_id)
            for therapist_id, count in counts.items()
            if count >= min_sessions
        ),
        key=lambda item: (-item[0], item[1]),
    )
    if not candidates:
        raise ValueError(
            f"No therapist has at least {min_sessions} scheduled sessions on the selected date"
        )

    _, therapist_id = candidates[0]
    affected = tuple(
        sorted(
            (session for session in eligible if session.therapist_id == therapist_id),
            key=lambda session: (session.start_time, session.patient_id, session.session_id),
        )
    )
    return therapist_id, affected


def _write_absence_row(workbook_path: Path, therapist_id: str) -> None:
    if sys.platform != "win32":
        raise RuntimeError("Native orphan smoke preparation requires Windows + Microsoft Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("pywin32 is required: python -m pip install pywin32") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)
        ws = workbook.Worksheets("DAILY_INPUT")
        ws.Range("A7").Value = therapist_id
        ws.Range("B7").Value = "ΝΑΙ"
        ws.Range("C7:D7").ClearContents()
        ws.Range("E7").Value = "Native orphan mixed smoke"
        ws.Range("F7").Value = "Resolve one by placement and one by Αναστολή θεραπείας"
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
        description=(
            "Create a safe DAILY_INPUT preview whose therapist absence produces at least "
            "two orphan sessions for the native placement+suspension smoke test."
        )
    )
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "DAILY_CHANGE_LEDGER_V85.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "ORPHAN_SMOKE_INPUT_V86.xlsm",
    )
    parser.add_argument("--min-sessions", type=int, default=2)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    source_hash_before = file_sha256(source)
    try:
        patient_audit = audit_daily_input_patients(source)
        if not patient_audit.ok:
            raise ValueError("Patient identity audit failed; smoke input was not created")

        settings = read_settings(source)
        base_entries = read_unified_base_schedule(source)
        sessions = materialize_sessions_for_date(base_entries, args.date)
        therapist_id, affected = choose_smoke_therapist(
            sessions,
            allowed_therapists=tuple(settings.therapist_names),
            min_sessions=args.min_sessions,
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
            source,
            output,
            spec,
            overwrite=args.overwrite,
        )
        _write_absence_row(preview, therapist_id)
    except (ValueError, DailyInputSheetError, RuntimeError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    source_unchanged = source_hash_before == file_sha256(source)
    vba_preserved = _has_vba(source) == _has_vba(output) and _has_vba(output)

    print("ORPHAN SMOKE INPUT OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Source: {source}")
    print(f"Preview: {output}")
    print(f"Selected absent therapist: {therapist_id}")
    print(f"Affected sessions: {len(affected)}")
    for session in affected:
        patient_name = patient_by_id.get(session.patient_id, session.patient_id)
        print(f"  {session.start_time.strftime('%H:%M')} | {patient_name} | {session.session_id}")
    print(f"Source unchanged: {source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print(
        "NEXT: run apply_daily_input_replacements.py on this preview. "
        "Choose a placement for at least one orphan and S for at least one other orphan."
    )
    return 0 if source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
