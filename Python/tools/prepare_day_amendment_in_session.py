from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import tempfile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.prepare_day_amendment import prepare_day_amendment  # noqa: E402


class DayAmendmentInSessionPreparationError(RuntimeError):
    pass


SNAPSHOT_FIELDS = (
    "snapshot_id",
    "service_date",
    "version",
    "status",
    "finalized_at",
    "supersedes_snapshot_id",
    "unresolved_sessions",
    "planned_session_count",
    "final_session_count",
    "patient_absence_count",
    "cancellation_count",
    "no_show_count",
    "replacement_count",
    "treatment_suspension_count",
    "approved_exception_count",
)

PATIENT_FIELDS = (
    "snapshot_id",
    "service_date",
    "snapshot_version",
    "session_id",
    "patient_id",
    "status",
    "original_therapist_id",
    "original_time",
    "effective_therapist_id",
    "effective_time",
    "replacement_id",
    "cancellation_id",
    "cancellation_kind",
    "reason",
)

THERAPIST_FIELDS = (
    "snapshot_id",
    "service_date",
    "therapist_id",
    "capacity_timeslots",
    "planned_occupied_timeslots",
    "planned_sessions",
    "final_occupied_timeslots",
    "final_active_sessions",
    "replacements_received",
    "released_by_cancellation_or_no_show",
    "treatment_suspensions_affecting_load",
    "approved_shared_or_over_capacity_events",
)

EVENT_FIELDS = (
    "event_id",
    "service_date",
    "event_type",
    "recorded_at",
    "actor_or_source",
    "therapist_id",
    "patient_id",
    "session_id",
    "reason_code",
    "related_event_id",
    "related_snapshot_id",
    "summary",
)


def _field(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    if "\t" in text or "\r" in text or "\n" in text:
        raise DayAmendmentInSessionPreparationError("Payload field contains unsupported control characters")
    return text


def _record(kind: str, data: dict[str, object], fields: tuple[str, ...]) -> str:
    return "\t".join((kind, *(_field(data.get(name)) for name in fields)))


def prepare_day_amendment_in_session(
    workbook_path: str | Path,
    payload_path: str | Path,
    *,
    actor_or_source: str,
) -> dict[str, object]:
    source = Path(workbook_path).resolve()
    payload = Path(payload_path).resolve()
    payload.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(
        prefix="eka_day_amendment_plan_",
        suffix=".json",
        dir=payload.parent,
        delete=False,
    ) as handle:
        json_path = Path(handle.name)

    try:
        result = prepare_day_amendment(
            source,
            json_path,
            actor_or_source=actor_or_source,
        )
        data = json.loads(json_path.read_text(encoding="utf-8-sig"))

        lines = [
            _record("SUPERSEDE", data["superseded_snapshot"], SNAPSHOT_FIELDS),
            _record("SNAPSHOT", data["replacement_snapshot"], SNAPSHOT_FIELDS),
        ]
        lines.extend(
            _record("PATIENT", row, PATIENT_FIELDS)
            for row in data["patient_rows"]
        )
        lines.extend(
            _record("THERAPIST", row, THERAPIST_FIELDS)
            for row in data["therapist_rows"]
        )
        lines.append(_record("EVENT", data["started_event"], EVENT_FIELDS))
        lines.append(_record("EVENT", data["finalized_event"], EVENT_FIELDS))
        payload.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    finally:
        try:
            json_path.unlink()
        except OSError:
            pass

    return {
        **result,
        "payload_path": str(payload),
        "payload_format": "day-amendment-in-session-tsv-v2",
        "event_rows": 2,
    }


def _write_response(path: Path | None, data: dict[str, object]) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8-sig")


def main() -> int:
    parser = argparse.ArgumentParser(description="Prepare an in-session day-save TSV payload.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--payload", required=True, type=Path)
    parser.add_argument("--response", type=Path)
    parser.add_argument("--actor-or-source", default="excel:day-save")
    args = parser.parse_args()

    try:
        result = prepare_day_amendment_in_session(
            args.input,
            args.payload,
            actor_or_source=args.actor_or_source,
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
    print("DAY SAVE IN-SESSION PREPARATION VERIFIED")
    print(f"Date: {result['service_date']}")
    print(f"Version: {result['previous_version']} -> {result['replacement_version']}")
    print(f"Patient/session history rows: {result['patient_history_rows']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
