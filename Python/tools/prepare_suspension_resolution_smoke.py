from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_excel.application_contract import (  # noqa: E402
    ApplicationContractError,
    assert_application_contract,
)
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_patients import audit_daily_input_patients  # noqa: E402
from rehab_excel.daily_input_reader import DailyInputReadError, read_daily_input  # noqa: E402
from rehab_excel.daily_input_sheet import (  # noqa: E402
    DailyInputSheetError,
    build_daily_input_spec,
    create_daily_input_preview,
)
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402


class SuspensionResolutionSmokeError(RuntimeError):
    pass


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def choose_robotic_smoke_target(sessions, *, allowed_therapists: tuple[str, ...]):
    allowed = set(allowed_therapists)
    robotic = [
        session
        for session in sessions
        if session.therapist_id in allowed and session.robotic
    ]
    if not robotic:
        raise ValueError("No robotic session is available on the smoke date")

    robotic.sort(key=lambda item: (item.therapist_id, item.start_time, item.patient_id, item.session_id))
    target = robotic[0]
    affected = tuple(
        sorted(
            (session for session in sessions if session.therapist_id == target.therapist_id),
            key=lambda item: (item.start_time, item.patient_id, item.session_id),
        )
    )
    return target.therapist_id, target, affected


def _write_smoke_rows(workbook_path: Path, *, absent_therapist_id: str) -> None:
    if sys.platform != "win32":
        raise SuspensionResolutionSmokeError("Windows + Microsoft Excel are required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise SuspensionResolutionSmokeError("pywin32 is required") from exc

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

        daily_ws = workbook.Worksheets("DAILY_INPUT")
        daily_ws.Cells(7, 1).Value = absent_therapist_id
        daily_ws.Cells(7, 2).Value = "ΝΑΙ"
        daily_ws.Range(daily_ws.Cells(7, 3), daily_ws.Cells(7, 4)).ClearContents()
        daily_ws.Cells(7, 5).Value = "Suspension-resolution smoke"
        daily_ws.Cells(7, 6).Value = "Robotic displaced session must become Αναστολή θεραπείας if unplaceable"

        # Keep the smoke deterministic even if the source later gains students:
        # no registered student in this disposable copy may satisfy robotic capability.
        try:
            students_ws = workbook.Worksheets("STUDENTS")
            last_row = students_ws.Cells(students_ws.Rows.Count, 1).End(-4162).Row
            for row in range(2, max(2, last_row) + 1):
                if str(students_ws.Cells(row, 1).Value or "").strip():
                    students_ws.Cells(row, 7).Value = "ΟΧΙ"
        except Exception:
            pass

        workbook.Save()
    except SuspensionResolutionSmokeError:
        raise
    except Exception as exc:
        raise SuspensionResolutionSmokeError(
            f"Excel suspension smoke preparation failed: {exc}"
        ) from exc
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
            "Build a full-app smoke copy with a robotic therapist-absence session "
            "that must finalize as explicit treatment suspension if unplaceable."
        )
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
        default=REPO_ROOT / "Excel" / "previews" / "SUSPENSION_RESOLUTION_INPUT.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    source_hash = file_sha256(source)
    output_created = False
    try:
        source_contract = assert_application_contract(source)
        patient_audit = audit_daily_input_patients(source)
        if not patient_audit.ok:
            raise SuspensionResolutionSmokeError("Patient identity audit failed")

        settings = read_settings(source)
        base_entries = read_unified_base_schedule(source)
        sessions = materialize_sessions_for_date(base_entries, args.date)
        absent_therapist_id, robotic_target, affected = choose_robotic_smoke_target(
            sessions,
            allowed_therapists=tuple(settings.therapist_names),
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
        output_created = True
        _write_smoke_rows(preview, absent_therapist_id=absent_therapist_id)

        daily = read_daily_input(
            preview,
            patients=patients,
            sessions=sessions,
            therapist_names=settings.therapist_names,
        )
        if not any(
            absence.subject_id == absent_therapist_id for absence in daily.absences
        ):
            raise SuspensionResolutionSmokeError("Therapist absence did not read back")
        output_contract = assert_application_contract(preview)
    except (
        ValueError,
        ApplicationContractError,
        DailyInputSheetError,
        DailyInputReadError,
        SuspensionResolutionSmokeError,
    ) as exc:
        if output_created:
            try:
                output.unlink(missing_ok=True)
            except OSError:
                pass
        print(f"SAFETY STOP: {exc}")
        return 2

    source_unchanged = file_sha256(source) == source_hash
    print("SUSPENSION RESOLUTION SMOKE INPUT OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Absent therapist: {absent_therapist_id}")
    print(
        "Forced robotic target: "
        f"{robotic_target.start_time.strftime('%H:%M')} | "
        f"{patient_by_id.get(robotic_target.patient_id, robotic_target.patient_id)}"
    )
    print(f"Affected sessions: {len(affected)}")
    print(f"Preview: {output}")
    print(f"Source unchanged: {source_unchanged}")
    print(f"Application contract source: OK ({len(source_contract.component_names)} components)")
    print(f"Application contract output: OK ({len(output_contract.component_names)} components)")
    print("NEXT: run apply_global_daily_replacements.py on this smoke input.")
    return 0 if source_unchanged else 3


if __name__ == "__main__":
    raise SystemExit(main())
