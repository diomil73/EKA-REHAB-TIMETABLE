from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.day_amendment import DayAmendmentPlan  # noqa: E402
from rehab_core.day_finalization import (  # noqa: E402
    FinalizedDaySnapshot,
    FinalizedSnapshotStatus,
)
from rehab_core.day_history_events import (  # noqa: E402
    DayHistoryEvent,
    DayHistoryEventType,
)
from rehab_core.patient_session_history import PatientSessionHistoryRow  # noqa: E402
from rehab_core.therapist_day_history import TherapistDayHistoryRow  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.day_amendment_store import persist_day_amendment  # noqa: E402
from rehab_excel.day_history_event_store import read_day_history_events  # noqa: E402
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots  # noqa: E402
from rehab_excel.patient_session_history_store import (  # noqa: E402
    _row_from_values as patient_row_from_values,
    read_patient_session_history_rows,
)
from rehab_excel.therapist_day_history_store import read_therapist_day_history_rows  # noqa: E402


class DayAmendmentApplyError(RuntimeError):
    pass


def _required_text(value: object, *, field: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise DayAmendmentApplyError(f"{field} is required")
    return text


def _parse_date(value: object, *, field: str) -> date:
    try:
        return date.fromisoformat(_required_text(value, field=field))
    except ValueError as exc:
        raise DayAmendmentApplyError(f"Invalid {field}: {value!r}") from exc


def _parse_datetime(value: object, *, field: str) -> datetime:
    try:
        parsed = datetime.fromisoformat(_required_text(value, field=field))
    except ValueError as exc:
        raise DayAmendmentApplyError(f"Invalid {field}: {value!r}") from exc
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        raise DayAmendmentApplyError(f"{field} must include a timezone offset")
    return parsed


def _int(value: object, *, field: str) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise DayAmendmentApplyError(f"Invalid {field}: {value!r}") from exc
    if result < 0:
        raise DayAmendmentApplyError(f"{field} cannot be negative")
    return result


def _optional_int(value: object, *, field: str) -> int | None:
    if value is None or str(value).strip() == "":
        return None
    return _int(value, field=field)


def _snapshot(data: dict[str, object], *, label: str) -> FinalizedDaySnapshot:
    try:
        status = FinalizedSnapshotStatus(_required_text(data.get("status"), field=f"{label}.status"))
    except ValueError as exc:
        raise DayAmendmentApplyError(f"Invalid {label}.status") from exc
    return FinalizedDaySnapshot(
        snapshot_id=_required_text(data.get("snapshot_id"), field=f"{label}.snapshot_id"),
        service_date=_parse_date(data.get("service_date"), field=f"{label}.service_date"),
        version=_int(data.get("version"), field=f"{label}.version"),
        status=status,
        finalized_at=_parse_datetime(data.get("finalized_at"), field=f"{label}.finalized_at"),
        supersedes_snapshot_id=(str(data.get("supersedes_snapshot_id") or "").strip() or None),
        unresolved_sessions=_int(data.get("unresolved_sessions"), field=f"{label}.unresolved_sessions"),
        planned_session_count=_optional_int(data.get("planned_session_count"), field=f"{label}.planned_session_count"),
        final_session_count=_int(data.get("final_session_count"), field=f"{label}.final_session_count"),
        patient_absence_count=_int(data.get("patient_absence_count"), field=f"{label}.patient_absence_count"),
        cancellation_count=_int(data.get("cancellation_count"), field=f"{label}.cancellation_count"),
        no_show_count=_int(data.get("no_show_count"), field=f"{label}.no_show_count"),
        replacement_count=_int(data.get("replacement_count"), field=f"{label}.replacement_count"),
        treatment_suspension_count=_int(data.get("treatment_suspension_count"), field=f"{label}.treatment_suspension_count"),
        approved_exception_count=_int(data.get("approved_exception_count"), field=f"{label}.approved_exception_count"),
    )


def _event(data: dict[str, object], *, label: str) -> DayHistoryEvent:
    try:
        event_type = DayHistoryEventType(_required_text(data.get("event_type"), field=f"{label}.event_type"))
    except ValueError as exc:
        raise DayAmendmentApplyError(f"Invalid {label}.event_type") from exc
    return DayHistoryEvent(
        event_id=_required_text(data.get("event_id"), field=f"{label}.event_id"),
        service_date=_parse_date(data.get("service_date"), field=f"{label}.service_date"),
        event_type=event_type,
        recorded_at=_parse_datetime(data.get("recorded_at"), field=f"{label}.recorded_at"),
        actor_or_source=_required_text(data.get("actor_or_source"), field=f"{label}.actor_or_source"),
        therapist_id=str(data.get("therapist_id") or "").strip() or None,
        patient_id=str(data.get("patient_id") or "").strip() or None,
        session_id=str(data.get("session_id") or "").strip() or None,
        reason_code=str(data.get("reason_code") or "").strip() or None,
        related_event_id=str(data.get("related_event_id") or "").strip() or None,
        related_snapshot_id=str(data.get("related_snapshot_id") or "").strip() or None,
        summary=str(data.get("summary") or "").strip(),
    )


def _therapist_row(data: dict[str, object]) -> TherapistDayHistoryRow:
    return TherapistDayHistoryRow(
        snapshot_id=_required_text(data.get("snapshot_id"), field="therapist_row.snapshot_id"),
        service_date=_parse_date(data.get("service_date"), field="therapist_row.service_date"),
        therapist_id=_required_text(data.get("therapist_id"), field="therapist_row.therapist_id"),
        capacity_timeslots=_optional_int(data.get("capacity_timeslots"), field="therapist_row.capacity_timeslots"),
        planned_occupied_timeslots=_int(data.get("planned_occupied_timeslots"), field="therapist_row.planned_occupied_timeslots"),
        planned_sessions=_int(data.get("planned_sessions"), field="therapist_row.planned_sessions"),
        final_occupied_timeslots=_int(data.get("final_occupied_timeslots"), field="therapist_row.final_occupied_timeslots"),
        final_active_sessions=_int(data.get("final_active_sessions"), field="therapist_row.final_active_sessions"),
        replacements_received=_int(data.get("replacements_received"), field="therapist_row.replacements_received"),
        released_by_cancellation_or_no_show=_int(data.get("released_by_cancellation_or_no_show"), field="therapist_row.released_by_cancellation_or_no_show"),
        treatment_suspensions_affecting_load=_int(data.get("treatment_suspensions_affecting_load"), field="therapist_row.treatment_suspensions_affecting_load"),
        approved_shared_or_over_capacity_events=_int(data.get("approved_shared_or_over_capacity_events"), field="therapist_row.approved_shared_or_over_capacity_events"),
    )


def _patient_row(data: dict[str, object]) -> PatientSessionHistoryRow:
    values = (
        data.get("snapshot_id"),
        data.get("service_date"),
        data.get("snapshot_version"),
        data.get("session_id"),
        data.get("patient_id"),
        data.get("status"),
        data.get("original_therapist_id"),
        data.get("original_time"),
        data.get("effective_therapist_id"),
        data.get("effective_time"),
        data.get("replacement_id"),
        data.get("cancellation_id"),
        data.get("cancellation_kind"),
        data.get("reason"),
    )
    try:
        return patient_row_from_values(values, row=0)
    except Exception as exc:
        raise DayAmendmentApplyError(f"Invalid patient_row: {exc}") from exc


def _dict(value: object, *, field: str) -> dict[str, object]:
    if not isinstance(value, dict):
        raise DayAmendmentApplyError(f"{field} must be an object")
    return value


def apply_day_amendment(workbook_path: str | Path, payload_path: str | Path) -> dict[str, object]:
    source = Path(workbook_path).resolve()
    payload = Path(payload_path).resolve()
    if not source.exists():
        raise DayAmendmentApplyError(f"Workbook not found: {source}")
    if not payload.exists():
        raise DayAmendmentApplyError(f"Payload not found: {payload}")

    try:
        data = json.loads(payload.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise DayAmendmentApplyError(f"Could not read amendment payload: {exc}") from exc
    if not isinstance(data, dict):
        raise DayAmendmentApplyError("Amendment payload root must be an object")
    schema = data.get("schema")
    if schema not in {"day-amendment-v1", "day-amendment-v2"}:
        raise DayAmendmentApplyError(f"Unsupported amendment payload schema: {schema!r}")

    payload_source = Path(_required_text(data.get("source_path"), field="source_path")).resolve()
    if payload_source != source:
        raise DayAmendmentApplyError(
            f"Payload source does not match target workbook: {payload_source} != {source}"
        )
    expected_hash = _required_text(data.get("source_sha256"), field="source_sha256")
    actual_hash = file_sha256(source)
    if actual_hash != expected_hash:
        raise DayAmendmentApplyError(
            "Workbook changed after amendment preparation; refusing to apply payload"
        )

    previous = _snapshot(_dict(data.get("previous_snapshot"), field="previous_snapshot"), label="previous_snapshot")
    superseded = _snapshot(_dict(data.get("superseded_snapshot"), field="superseded_snapshot"), label="superseded_snapshot")
    replacement = _snapshot(_dict(data.get("replacement_snapshot"), field="replacement_snapshot"), label="replacement_snapshot")
    started = _event(_dict(data.get("started_event"), field="started_event"), label="started_event")
    finalized = _event(_dict(data.get("finalized_event"), field="finalized_event"), label="finalized_event")

    raw_rows = data.get("therapist_rows")
    if not isinstance(raw_rows, list):
        raise DayAmendmentApplyError("therapist_rows must be an array")
    therapist_rows = tuple(_therapist_row(_dict(item, field="therapist_row")) for item in raw_rows)

    patient_rows: tuple[PatientSessionHistoryRow, ...] | None = None
    if schema == "day-amendment-v2":
        raw_patient_rows = data.get("patient_rows")
        if not isinstance(raw_patient_rows, list):
            raise DayAmendmentApplyError("patient_rows must be an array for day-amendment-v2")
        patient_rows = tuple(
            _patient_row(_dict(item, field="patient_row")) for item in raw_patient_rows
        )

    plan = DayAmendmentPlan(
        previous_snapshot=previous,
        superseded_snapshot=superseded,
        replacement_snapshot=replacement,
        started_event=started,
        finalized_event=finalized,
    )
    persisted = persist_day_amendment(
        source,
        plan,
        therapist_rows,
        patient_rows=patient_rows,
    )

    snapshots = read_finalized_day_snapshots(source, service_date=replacement.service_date)
    by_id = {item.snapshot_id: item for item in snapshots}
    if by_id.get(previous.snapshot_id) != superseded:
        raise DayAmendmentApplyError("Post-save verification failed for superseded snapshot")
    if by_id.get(replacement.snapshot_id) != replacement:
        raise DayAmendmentApplyError("Post-save verification failed for replacement snapshot")
    current = [item for item in snapshots if item.status == FinalizedSnapshotStatus.CURRENT_FINAL]
    if current != [replacement]:
        raise DayAmendmentApplyError("Post-save verification did not find exactly the replacement CURRENT_FINAL snapshot")

    if patient_rows is not None:
        persisted_patient_rows = read_patient_session_history_rows(
            source,
            snapshot_id=replacement.snapshot_id,
        )
        if persisted_patient_rows != patient_rows:
            raise DayAmendmentApplyError("Post-save patient/session history verification failed")

    persisted_rows = read_therapist_day_history_rows(source, snapshot_id=replacement.snapshot_id)
    if persisted_rows != therapist_rows:
        raise DayAmendmentApplyError("Post-save therapist/day history verification failed")

    event_ids = {started.event_id, finalized.event_id}
    persisted_events = tuple(
        event for event in read_day_history_events(source, service_date=replacement.service_date)
        if event.event_id in event_ids
    )
    if {event.event_id for event in persisted_events} != event_ids:
        raise DayAmendmentApplyError("Post-save amendment event verification failed")

    return {
        "ok": True,
        **persisted,
        "source_sha256_before": expected_hash,
        "source_sha256_after": file_sha256(source),
        "post_save_verified": True,
    }


def _write_response(path: Path | None, data: dict[str, object]) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8-sig")


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply one prepared closed-day amendment payload.")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--payload", required=True, type=Path)
    parser.add_argument("--response", type=Path)
    args = parser.parse_args()

    try:
        result = apply_day_amendment(args.workbook, args.payload)
    except Exception as exc:
        _write_response(args.response, {"ok": False, "error": str(exc)})
        print(f"SAFETY STOP: {exc}")
        return 2

    _write_response(args.response, result)
    print("DAY AMENDMENT APPLY VERIFIED")
    print(f"Date: {result['service_date']}")
    print(f"Replacement snapshot: {result['replacement_snapshot_id']}")
    print(f"Version: {result['replacement_version']}")
    if result.get("patient_history_rows") is not None:
        print(f"Patient/session history rows: {result['patient_history_rows']}")
    print(f"Therapist/day history rows: {result['therapist_history_rows']}")
    print("Post-save verification: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
