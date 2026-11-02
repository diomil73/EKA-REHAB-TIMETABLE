from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from pathlib import Path
import sys

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.day_amendment import build_day_amendment_plan  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_core.therapist_day_history import build_therapist_day_history_rows  # noqa: E402
from rehab_excel.application_contract import assert_application_contract  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_reader import read_daily_input  # noqa: E402
from rehab_excel.daily_replacement_facts import read_daily_replacement_facts  # noqa: E402
from rehab_excel.day_history_event_store import read_day_history_events  # noqa: E402
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots  # noqa: E402
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402
from rehab_excel.therapist_day_history_store import read_therapist_day_history_rows  # noqa: E402


AUTO_REASON_CODE = "AUTO_DAY_SAVE"
AUTO_SUMMARY = "Automatic day save from workbook state"


class DayAmendmentPreparationError(RuntimeError):
    pass


def _daily_input_date(path: Path) -> date:
    wb = load_workbook(
        path,
        read_only=True,
        data_only=False,
        keep_vba=path.suffix.casefold() == ".xlsm",
    )
    try:
        if "DAILY_INPUT" not in wb.sheetnames:
            raise DayAmendmentPreparationError("Workbook has no DAILY_INPUT sheet")
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
    raise DayAmendmentPreparationError(f"Invalid DAILY_INPUT date: {raw!r}")


def _is_legacy_student_name(value: str) -> bool:
    compact = str(value).strip().casefold().replace(" ", "")
    return compact.startswith("φοιτ") or compact.startswith("foit")


def _snapshot_dict(snapshot) -> dict[str, object]:
    return {
        "snapshot_id": snapshot.snapshot_id,
        "service_date": snapshot.service_date.isoformat(),
        "version": snapshot.version,
        "status": snapshot.status.value,
        "finalized_at": snapshot.finalized_at.isoformat(),
        "supersedes_snapshot_id": snapshot.supersedes_snapshot_id,
        "unresolved_sessions": snapshot.unresolved_sessions,
        "planned_session_count": snapshot.planned_session_count,
        "final_session_count": snapshot.final_session_count,
        "patient_absence_count": snapshot.patient_absence_count,
        "cancellation_count": snapshot.cancellation_count,
        "no_show_count": snapshot.no_show_count,
        "replacement_count": snapshot.replacement_count,
        "treatment_suspension_count": snapshot.treatment_suspension_count,
        "approved_exception_count": snapshot.approved_exception_count,
    }


def _event_dict(event) -> dict[str, object]:
    return {
        "event_id": event.event_id,
        "service_date": event.service_date.isoformat(),
        "event_type": event.event_type.value,
        "recorded_at": event.recorded_at.isoformat(),
        "actor_or_source": event.actor_or_source,
        "therapist_id": event.therapist_id,
        "patient_id": event.patient_id,
        "session_id": event.session_id,
        "reason_code": event.reason_code,
        "related_event_id": event.related_event_id,
        "related_snapshot_id": event.related_snapshot_id,
        "summary": event.summary,
    }


def _therapist_row_dict(row) -> dict[str, object]:
    return {
        "snapshot_id": row.snapshot_id,
        "service_date": row.service_date.isoformat(),
        "therapist_id": row.therapist_id,
        "capacity_timeslots": row.capacity_timeslots,
        "planned_occupied_timeslots": row.planned_occupied_timeslots,
        "planned_sessions": row.planned_sessions,
        "final_occupied_timeslots": row.final_occupied_timeslots,
        "final_active_sessions": row.final_active_sessions,
        "replacements_received": row.replacements_received,
        "released_by_cancellation_or_no_show": row.released_by_cancellation_or_no_show,
        "treatment_suspensions_affecting_load": row.treatment_suspensions_affecting_load,
        "approved_shared_or_over_capacity_events": row.approved_shared_or_over_capacity_events,
    }


def prepare_day_amendment(
    workbook_path: str | Path,
    payload_path: str | Path,
    *,
    actor_or_source: str,
    reason_code: str = AUTO_REASON_CODE,
    summary: str = AUTO_SUMMARY,
    approved_exception_count: int = 0,
) -> dict[str, object]:
    source = Path(workbook_path).resolve()
    payload = Path(payload_path).resolve()
    if not source.exists():
        raise DayAmendmentPreparationError(f"Workbook not found: {source}")
    if source.suffix.casefold() != ".xlsm":
        raise DayAmendmentPreparationError("Workbook must be an .xlsm file")
    if approved_exception_count < 0:
        raise DayAmendmentPreparationError("approved_exception_count cannot be negative")

    actor = actor_or_source.strip()
    reason = reason_code.strip()
    summary_text = summary.strip()
    if not actor:
        raise DayAmendmentPreparationError("actor_or_source cannot be blank")
    if not reason:
        raise DayAmendmentPreparationError("reason_code cannot be blank")
    if not summary_text:
        raise DayAmendmentPreparationError("summary cannot be blank")

    source_hash = file_sha256(source)
    contract = assert_application_contract(source)
    service_date = _daily_input_date(source)

    current = read_finalized_day_snapshots(
        source,
        service_date=service_date,
        current_only=True,
    )
    if len(current) != 1:
        raise DayAmendmentPreparationError(
            f"Expected exactly one CURRENT_FINAL snapshot for {service_date.isoformat()}, found {len(current)}"
        )
    previous = current[0]

    patients = read_patient_registry(source)
    settings = read_settings(source)
    base_entries = read_unified_base_schedule(source)
    sessions = materialize_sessions_for_date(base_entries, service_date)
    daily = read_daily_input(
        source,
        patients=patients,
        sessions=sessions,
        therapist_names=settings.therapist_names,
        allow_empty=True,
    )
    replacements = read_daily_replacement_facts(
        source,
        sessions=sessions,
        target_date=service_date,
    )
    states = build_daily_session_states(
        sessions,
        absences=daily.absences,
        replacements=replacements,
        cancellations=daily.cancellations,
        target_date=service_date,
    )

    started_at = datetime.now().astimezone()
    finalized_at = datetime.now().astimezone()
    next_version = previous.version + 1
    replacement_id = f"finalized:{service_date.isoformat()}:v{next_version}"
    started_event_id = f"day-amendment-started:{service_date.isoformat()}:v{next_version}"
    finalized_event_id = f"day-amendment-finalized:{service_date.isoformat()}:v{next_version}"

    plan = build_day_amendment_plan(
        states,
        previous_snapshot=previous,
        replacement_snapshot_id=replacement_id,
        started_event_id=started_event_id,
        finalized_event_id=finalized_event_id,
        actor_or_source=actor,
        amendment_started_at=started_at,
        amendment_finalized_at=finalized_at,
        reason_code=reason,
        summary=summary_text,
        planned_session_count=len(sessions),
        approved_exception_count=approved_exception_count,
    )

    therapist_ids = tuple(
        name
        for name in settings.therapist_names
        if str(name).strip() and not _is_legacy_student_name(name)
    )
    therapist_rows = build_therapist_day_history_rows(
        states,
        snapshot_id=plan.replacement_snapshot.snapshot_id,
        service_date=service_date,
        therapist_ids=therapist_ids,
        capacity_timeslots_by_therapist={name: 6 for name in therapist_ids},
    )

    if read_therapist_day_history_rows(source, snapshot_id=replacement_id):
        raise DayAmendmentPreparationError("Replacement therapist/day history already exists")
    existing_event_ids = {item.event_id for item in read_day_history_events(source)}
    collisions = {started_event_id, finalized_event_id} & existing_event_ids
    if collisions:
        raise DayAmendmentPreparationError(
            f"Amendment EventId already exists: {sorted(collisions)!r}"
        )

    payload_data = {
        "schema": "day-amendment-v1",
        "source_path": str(source),
        "source_sha256": source_hash,
        "service_date": service_date.isoformat(),
        "previous_snapshot": _snapshot_dict(plan.previous_snapshot),
        "superseded_snapshot": _snapshot_dict(plan.superseded_snapshot),
        "replacement_snapshot": _snapshot_dict(plan.replacement_snapshot),
        "therapist_rows": [_therapist_row_dict(row) for row in therapist_rows],
        "started_event": _event_dict(plan.started_event),
        "finalized_event": _event_dict(plan.finalized_event),
    }
    payload.parent.mkdir(parents=True, exist_ok=True)
    payload.write_text(
        json.dumps(payload_data, ensure_ascii=False, indent=2),
        encoding="utf-8-sig",
    )

    if file_sha256(source) != source_hash:
        try:
            payload.unlink()
        except OSError:
            pass
        raise DayAmendmentPreparationError("Workbook changed during amendment preparation")

    return {
        "ok": True,
        "payload_path": str(payload),
        "source_sha256": source_hash,
        "service_date": service_date.isoformat(),
        "previous_snapshot_id": previous.snapshot_id,
        "previous_version": previous.version,
        "replacement_snapshot_id": plan.replacement_snapshot.snapshot_id,
        "replacement_version": plan.replacement_snapshot.version,
        "therapist_history_rows": len(therapist_rows),
        "unresolved_sessions": plan.replacement_snapshot.unresolved_sessions,
        "application_contract_components": len(contract.component_names),
        "source_unchanged": True,
    }


def _write_response(path: Path | None, data: dict[str, object]) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8-sig")


def main() -> int:
    parser = argparse.ArgumentParser(description="Prepare a read-only closed-day amendment payload.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--payload", required=True, type=Path)
    parser.add_argument("--response", type=Path)
    parser.add_argument("--actor-or-source", default="day-amendment")
    parser.add_argument("--reason-code", default=AUTO_REASON_CODE)
    parser.add_argument("--summary", default=AUTO_SUMMARY)
    parser.add_argument("--approved-exception-count", type=int, default=0)
    args = parser.parse_args()

    try:
        result = prepare_day_amendment(
            args.input,
            args.payload,
            actor_or_source=args.actor_or_source,
            reason_code=args.reason_code,
            summary=args.summary,
            approved_exception_count=args.approved_exception_count,
        )
    except Exception as exc:
        try:
            if args.payload.exists():
                args.payload.unlink()
        except OSError:
            pass
        _write_response(args.response, {"ok": False, "error": str(exc)})
        print(f"SAFETY STOP: {exc}")
        return 2

    _write_response(args.response, result)
    print("DAY AMENDMENT PREPARATION VERIFIED")
    print(f"Date: {result['service_date']}")
    print(
        f"Version: {result['previous_version']} -> {result['replacement_version']}"
    )
    print(f"Replacement snapshot: {result['replacement_snapshot_id']}")
    print(f"Therapist/day history rows: {result['therapist_history_rows']}")
    print("Input unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
