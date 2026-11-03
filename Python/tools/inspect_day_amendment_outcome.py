from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.day_finalization import FinalizedSnapshotStatus  # noqa: E402
from rehab_core.day_history_events import DayHistoryEventType  # noqa: E402
from rehab_excel.day_history_event_store import read_day_history_events  # noqa: E402
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots  # noqa: E402
from rehab_excel.patient_session_history_store import (  # noqa: E402
    read_current_patient_session_history_rows,
    read_patient_session_history_rows,
)
from rehab_excel.therapist_day_history_store import read_therapist_day_history_rows  # noqa: E402


class DayAmendmentInspectionError(RuntimeError):
    pass


def inspect_day_amendment_outcome(workbook_path: str | Path) -> dict[str, object]:
    path = Path(workbook_path).resolve()
    if not path.exists():
        raise DayAmendmentInspectionError(f"Workbook not found: {path}")

    snapshots = read_finalized_day_snapshots(path)
    if not snapshots:
        raise DayAmendmentInspectionError("No finalized snapshots found")

    by_date: dict[object, list] = {}
    for snapshot in snapshots:
        by_date.setdefault(snapshot.service_date, []).append(snapshot)

    amended_dates = []
    for service_date, items in by_date.items():
        currents = [item for item in items if item.status == FinalizedSnapshotStatus.CURRENT_FINAL]
        superseded = [item for item in items if item.status == FinalizedSnapshotStatus.SUPERSEDED]
        if len(currents) == 1 and superseded:
            amended_dates.append(service_date)

    if len(amended_dates) != 1:
        raise DayAmendmentInspectionError(
            f"Expected exactly one amended service date, found {len(amended_dates)}"
        )

    service_date = amended_dates[0]
    items = sorted(by_date[service_date], key=lambda item: item.version)
    current = [item for item in items if item.status == FinalizedSnapshotStatus.CURRENT_FINAL][0]
    prior = [item for item in items if item.snapshot_id == current.supersedes_snapshot_id]
    if len(prior) != 1:
        raise DayAmendmentInspectionError("Current snapshot does not link to exactly one prior snapshot")
    previous = prior[0]
    if previous.status != FinalizedSnapshotStatus.SUPERSEDED:
        raise DayAmendmentInspectionError("Previous snapshot is not SUPERSEDED")
    if current.version != previous.version + 1:
        raise DayAmendmentInspectionError("Snapshot version chain is not contiguous")
    if current.unresolved_sessions != 0:
        raise DayAmendmentInspectionError("Current amended snapshot has unresolved sessions")

    therapist_rows = read_therapist_day_history_rows(path, snapshot_id=current.snapshot_id)
    if not therapist_rows:
        raise DayAmendmentInspectionError("No therapist/day history rows for amended snapshot")

    previous_patient_rows = read_patient_session_history_rows(
        path,
        snapshot_id=previous.snapshot_id,
    )
    current_patient_rows = read_patient_session_history_rows(
        path,
        snapshot_id=current.snapshot_id,
    )
    operational_patient_rows = read_current_patient_session_history_rows(
        path,
        service_date=service_date,
    )
    if not previous_patient_rows:
        raise DayAmendmentInspectionError("No patient/session history rows for superseded snapshot")
    if not current_patient_rows:
        raise DayAmendmentInspectionError("No patient/session history rows for current snapshot")
    if {row.snapshot_id for row in operational_patient_rows} != {current.snapshot_id}:
        raise DayAmendmentInspectionError("Operational patient history is not restricted to CURRENT_FINAL")
    if {(row.snapshot_id, row.session_id) for row in operational_patient_rows} != {
        (row.snapshot_id, row.session_id) for row in current_patient_rows
    }:
        raise DayAmendmentInspectionError("Operational patient history does not match current snapshot rows")

    events = read_day_history_events(path, service_date=service_date)
    started = [event for event in events if event.event_type == DayHistoryEventType.DAY_AMENDMENT_STARTED]
    finalized = [event for event in events if event.event_type == DayHistoryEventType.DAY_AMENDMENT_FINALIZED]
    if len(started) != 1 or len(finalized) != 1:
        raise DayAmendmentInspectionError(
            f"Expected one amendment started/finalized event, found {len(started)}/{len(finalized)}"
        )
    if started[0].related_snapshot_id != previous.snapshot_id:
        raise DayAmendmentInspectionError("Amendment-started event does not reference previous snapshot")
    if finalized[0].related_event_id != started[0].event_id:
        raise DayAmendmentInspectionError("Amendment-finalized event does not reference started event")
    if finalized[0].related_snapshot_id != current.snapshot_id:
        raise DayAmendmentInspectionError("Amendment-finalized event does not reference current snapshot")

    return {
        "ok": True,
        "service_date": service_date.isoformat(),
        "previous_snapshot_id": previous.snapshot_id,
        "previous_version": previous.version,
        "current_snapshot_id": current.snapshot_id,
        "current_version": current.version,
        "therapist_history_rows": len(therapist_rows),
        "previous_patient_history_rows": len(previous_patient_rows),
        "current_patient_history_rows": len(current_patient_rows),
        "operational_patient_history_rows": len(operational_patient_rows),
        "amendment_events": 2,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect a persisted closed-day amendment outcome.")
    parser.add_argument("--workbook", required=True, type=Path)
    args = parser.parse_args()

    try:
        result = inspect_day_amendment_outcome(args.workbook)
    except Exception as exc:
        print(f"DAY AMENDMENT INSPECTION FAILED: {exc}")
        return 2

    print("DAY AMENDMENT OUTCOME OK")
    print(f"Date: {result['service_date']}")
    print(f"Version: {result['previous_version']} -> {result['current_version']}")
    print(f"Current snapshot: {result['current_snapshot_id']}")
    print(f"Therapist/day history rows: {result['therapist_history_rows']}")
    print(f"Previous patient/session history rows: {result['previous_patient_history_rows']}")
    print(f"Current patient/session history rows: {result['current_patient_history_rows']}")
    print(f"Operational CURRENT_FINAL patient rows: {result['operational_patient_history_rows']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
