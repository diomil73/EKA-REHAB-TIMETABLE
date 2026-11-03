from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path
import sys

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.day_history_events import DayHistoryEventType  # noqa: E402
from rehab_excel.day_history_event_store import (  # noqa: E402
    SHEET_NAME as EVENT_SHEET,
    read_day_history_events,
)
from rehab_excel.finalized_day_snapshots import (  # noqa: E402
    SHEET_NAME as SNAPSHOT_SHEET,
    read_finalized_day_snapshots,
)
from rehab_excel.patient_session_history_store import (  # noqa: E402
    SHEET_NAME as PATIENT_SHEET,
    read_current_patient_session_history_rows,
    read_patient_session_history_rows,
)
from rehab_excel.therapist_day_history_store import (  # noqa: E402
    SHEET_NAME as THERAPIST_SHEET,
    read_therapist_day_history_rows,
)


class CloseDayInspectionError(RuntimeError):
    pass


def _service_date(path: Path) -> date:
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        if "DAILY_INPUT" not in wb.sheetnames:
            raise CloseDayInspectionError("Workbook has no DAILY_INPUT sheet")
        raw = wb["DAILY_INPUT"]["B2"].value
    finally:
        wb.close()

    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    text = str(raw or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise CloseDayInspectionError(f"Invalid DAILY_INPUT date: {raw!r}")


def _sheet_states(path: Path) -> dict[str, str]:
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        states: dict[str, str] = {}
        for name in (SNAPSHOT_SHEET, THERAPIST_SHEET, PATIENT_SHEET, EVENT_SHEET):
            if name not in wb.sheetnames:
                raise CloseDayInspectionError(f"Missing history sheet: {name}")
            states[name] = wb[name].sheet_state
        return states
    finally:
        wb.close()


def inspect_close_day_outcome(workbook_path: str | Path) -> dict[str, object]:
    path = Path(workbook_path).resolve()
    if not path.exists():
        raise CloseDayInspectionError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise CloseDayInspectionError("Workbook must be an .xlsm file")

    service_date = _service_date(path)
    snapshots = read_finalized_day_snapshots(
        path,
        service_date=service_date,
        current_only=True,
    )
    if len(snapshots) != 1:
        raise CloseDayInspectionError(
            f"Expected exactly one CURRENT_FINAL snapshot for {service_date.isoformat()}, found {len(snapshots)}"
        )
    snapshot = snapshots[0]
    if snapshot.unresolved_sessions != 0:
        raise CloseDayInspectionError("Finalized snapshot still has unresolved sessions")

    therapist_rows = read_therapist_day_history_rows(
        path,
        snapshot_id=snapshot.snapshot_id,
        service_date=service_date,
    )
    if not therapist_rows:
        raise CloseDayInspectionError("No therapist/day history rows for finalized snapshot")
    if any(row.snapshot_id != snapshot.snapshot_id for row in therapist_rows):
        raise CloseDayInspectionError("Therapist/day history references another snapshot")
    if any(row.service_date != service_date for row in therapist_rows):
        raise CloseDayInspectionError("Therapist/day history references another service date")

    patient_rows = read_patient_session_history_rows(
        path,
        snapshot_id=snapshot.snapshot_id,
        service_date=service_date,
    )
    if not patient_rows:
        raise CloseDayInspectionError("No patient/session history rows for finalized snapshot")
    if any(row.snapshot_id != snapshot.snapshot_id for row in patient_rows):
        raise CloseDayInspectionError("Patient/session history references another snapshot")
    if any(row.service_date != service_date for row in patient_rows):
        raise CloseDayInspectionError("Patient/session history references another service date")
    if any(row.snapshot_version != snapshot.version for row in patient_rows):
        raise CloseDayInspectionError("Patient/session history references another snapshot version")

    current_patient_rows = read_current_patient_session_history_rows(
        path,
        service_date=service_date,
    )
    patient_keys = {(row.snapshot_id, row.session_id) for row in patient_rows}
    current_patient_keys = {(row.snapshot_id, row.session_id) for row in current_patient_rows}
    if current_patient_keys != patient_keys:
        raise CloseDayInspectionError(
            "CURRENT_FINAL patient/session history does not exactly match finalized snapshot rows"
        )

    finalized_events = read_day_history_events(
        path,
        service_date=service_date,
        event_type=DayHistoryEventType.DAY_FINALIZED,
    )
    matching_events = tuple(
        event for event in finalized_events
        if event.related_snapshot_id == snapshot.snapshot_id
    )
    if len(matching_events) != 1:
        raise CloseDayInspectionError(
            f"Expected exactly one DAY_FINALIZED event linked to {snapshot.snapshot_id}, found {len(matching_events)}"
        )
    event = matching_events[0]

    states = _sheet_states(path)
    not_very_hidden = {
        name: state for name, state in states.items() if state != "veryHidden"
    }
    if not_very_hidden:
        raise CloseDayInspectionError(
            f"History sheets are not VeryHidden: {not_very_hidden!r}"
        )

    return {
        "workbook": str(path),
        "service_date": service_date.isoformat(),
        "snapshot_id": snapshot.snapshot_id,
        "snapshot_version": snapshot.version,
        "snapshot_status": snapshot.status.value,
        "finalized_at": snapshot.finalized_at.isoformat(),
        "planned_sessions": snapshot.planned_session_count,
        "final_sessions": snapshot.final_session_count,
        "patient_absences": snapshot.patient_absence_count,
        "cancellations": snapshot.cancellation_count,
        "no_shows": snapshot.no_show_count,
        "replacements": snapshot.replacement_count,
        "treatment_suspensions": snapshot.treatment_suspension_count,
        "approved_exceptions": snapshot.approved_exception_count,
        "therapist_history_rows": len(therapist_rows),
        "patient_history_rows": len(patient_rows),
        "event_id": event.event_id,
        "event_actor": event.actor_or_source,
        "event_summary": event.summary,
        "history_sheet_states": states,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read-only inspection of one native Close Day outcome."
    )
    parser.add_argument("--workbook", required=True, type=Path)
    args = parser.parse_args()

    try:
        result = inspect_close_day_outcome(args.workbook)
    except Exception as exc:
        print(f"CLOSE DAY INSPECTION FAILED: {exc}")
        return 2

    print("CLOSE DAY NATIVE OUTCOME OK")
    print(f"WORKBOOK: {result['workbook']}")
    print(f"SERVICE_DATE: {result['service_date']}")
    print(
        "SNAPSHOT: "
        f"id={result['snapshot_id']} version={result['snapshot_version']} "
        f"status={result['snapshot_status']} finalized_at={result['finalized_at']}"
    )
    print(
        "COUNTS: "
        f"planned={result['planned_sessions']} final={result['final_sessions']} "
        f"patient_absences={result['patient_absences']} cancellations={result['cancellations']} "
        f"no_shows={result['no_shows']} replacements={result['replacements']} "
        f"treatment_suspensions={result['treatment_suspensions']} "
        f"approved_exceptions={result['approved_exceptions']}"
    )
    print(f"THERAPIST_HISTORY_ROWS: {result['therapist_history_rows']}")
    print(f"PATIENT_HISTORY_ROWS: {result['patient_history_rows']}")
    print(
        "DAY_FINALIZED_EVENT: "
        f"id={result['event_id']} actor={result['event_actor']} "
        f"summary={result['event_summary']!r}"
    )
    for name, state in result["history_sheet_states"].items():
        print(f"SHEET_STATE {name}: {state}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
