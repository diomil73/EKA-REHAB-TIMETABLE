from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.models import AbsenceKind  # noqa: E402
from rehab_excel.daily_input_reader import DailyInputReadError, read_daily_input  # noqa: E402
from rehab_excel.reader import read_base_schedule, read_patients  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from datetime import date, datetime  # noqa: E402


def _read_date(path: Path) -> date:
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        value = wb["DAILY_INPUT"]["B2"].value
    finally:
        wb.close()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise DailyInputReadError(f"Invalid DAILY_INPUT date: {value!r}")


def main() -> int:
    path = (REPO_ROOT / "Excel" / "previews" / "DAILY_INPUT_PREVIEW.xlsm").resolve()
    if not path.exists():
        print(f"DIAGNOSTIC STOP: workbook not found: {path}")
        return 2
    try:
        patients = read_patients(path)
        entries = read_base_schedule(path)
        target_date = _read_date(path)
        sessions = materialize_sessions_for_date(entries, target_date)
        daily = read_daily_input(path, patients=patients, sessions=sessions)
    except Exception as exc:
        print(f"DIAGNOSTIC STOP: {exc}")
        return 2

    patient_by_id = {p.patient_id: p for p in patients}
    print("DAILY INPUT MATCH DIAGNOSTIC")
    print(f"Workbook: {path}")
    print(f"Date: {target_date.isoformat()} ({target_date.strftime('%A')})")
    print(f"Base entries: {len(entries)}")
    print(f"Sessions on selected date: {len(sessions)}")
    print(f"Absence rows parsed: {len(daily.absences)}")

    total_matches = 0
    for idx, absence in enumerate(daily.absences, start=1):
        if absence.absence_kind == AbsenceKind.PATIENT:
            patient = patient_by_id.get(absence.subject_id)
            label = patient.display_name if patient else absence.subject_id
            candidates = [s for s in sessions if s.patient_id == absence.subject_id]
            matches = [s for s in candidates if absence.covers(s.session_date, s.start_time)]
            print(f"  {idx}. PATIENT {label} id={absence.subject_id}")
            print(
                "     input window: "
                + ("ALL DAY" if absence.start_time is None else f"{absence.start_time.strftime('%H:%M')}"
                   + (f"-{absence.end_time.strftime('%H:%M')}" if absence.end_time else ""))
            )
            if candidates:
                print("     scheduled today: " + ", ".join(
                    f"{s.therapist_id}@{s.start_time.strftime('%H:%M')} ({s.treatment or ''})" for s in candidates
                ))
            else:
                print("     scheduled today: NONE")
            print(f"     matching sessions: {len(matches)}")
            total_matches += len(matches)
        else:
            candidates = [s for s in sessions if s.therapist_id == absence.subject_id]
            matches = [s for s in candidates if absence.covers(s.session_date, s.start_time)]
            print(f"  {idx}. THERAPIST {absence.subject_id}")
            print(
                "     input window: "
                + ("ALL DAY" if absence.start_time is None else f"{absence.start_time.strftime('%H:%M')}"
                   + (f"-{absence.end_time.strftime('%H:%M')}" if absence.end_time else ""))
            )
            if candidates:
                print("     scheduled today: " + ", ".join(
                    f"{patient_by_id.get(s.patient_id).display_name if patient_by_id.get(s.patient_id) else s.patient_id}@{s.start_time.strftime('%H:%M')}" for s in candidates
                ))
            else:
                print("     scheduled today: NONE")
            print(f"     matching sessions: {len(matches)}")
            total_matches += len(matches)

    print(f"TOTAL MATCHING SESSIONS: {total_matches}")
    if total_matches == 0:
        print("RESULT: the DAILY_INPUT row was parsed, but it does not intersect the materialized schedule for this date/time.")
        return 3
    print("RESULT: at least one DAILY_INPUT row matches a scheduled session.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
