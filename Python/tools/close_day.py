from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.day_history_events import (  # noqa: E402
    DayHistoryEventError,
    DayHistoryEventType,
    build_day_finalized_event,
)
from rehab_excel.application_contract import (  # noqa: E402
    ApplicationContractError,
    assert_application_contract,
)
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.day_history_event_store import (  # noqa: E402
    DayHistoryEventStoreError,
    append_day_history_events,
    read_day_history_events,
)
from rehab_excel.therapist_day_history_store import read_therapist_day_history_rows  # noqa: E402
from tools.build_finalized_day_snapshot_preview import (  # type: ignore[import-not-found] # noqa: E402
    FinalizedDayPreviewError,
    build_finalized_day_snapshot_preview,
)


class CloseDayError(RuntimeError):
    pass


def _write_response(path: Path | None, payload: dict[str, object]) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8-sig",
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Validate one operational day, create a verified finalized-history preview, "
            "and append the DAY_FINALIZED audit event."
        )
    )
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--response", type=Path)
    parser.add_argument("--actor-or-source", default="close-day-action")
    parser.add_argument("--approved-exception-count", type=int, default=0)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.input.resolve()
    output = args.output.resolve()
    if not source.exists():
        message = f"input workbook not found: {source}"
        _write_response(args.response, {"ok": False, "error": message})
        print(f"SAFETY STOP: {message}")
        return 2

    source_sha256_before = file_sha256(source)
    try:
        snapshot, source_components, _ = build_finalized_day_snapshot_preview(
            source,
            output,
            approved_exception_count=args.approved_exception_count,
            overwrite=args.overwrite,
        )
        event = build_day_finalized_event(
            event_id=f"day-finalized:{snapshot.service_date.isoformat()}:v{snapshot.version}",
            service_date=snapshot.service_date,
            recorded_at=snapshot.finalized_at,
            actor_or_source=args.actor_or_source,
            snapshot_id=snapshot.snapshot_id,
            summary=(
                f"Day finalized: {snapshot.final_session_count} performed sessions, "
                f"{snapshot.replacement_count} replacements, "
                f"{snapshot.treatment_suspension_count} suspensions"
            ),
        )
        if append_day_history_events(output, (event,)) != 1:
            raise CloseDayError("DAY_FINALIZED event append count was not 1")
        reread = read_day_history_events(
            output,
            service_date=snapshot.service_date,
            event_type=DayHistoryEventType.DAY_FINALIZED,
        )
        if event not in reread:
            raise CloseDayError("DAY_FINALIZED event did not read back exactly")
        therapist_rows = read_therapist_day_history_rows(
            output,
            snapshot_id=snapshot.snapshot_id,
            service_date=snapshot.service_date,
        )
        output_contract = assert_application_contract(output)
        if file_sha256(source) != source_sha256_before:
            raise CloseDayError("Input workbook changed during close-day preview build")
    except (
        CloseDayError,
        FinalizedDayPreviewError,
        DayHistoryEventError,
        DayHistoryEventStoreError,
        ApplicationContractError,
        OSError,
        ValueError,
        KeyError,
    ) as exc:
        try:
            if output.exists():
                output.unlink()
        except Exception:
            pass
        _write_response(args.response, {"ok": False, "error": str(exc)})
        print(f"SAFETY STOP: {exc}")
        return 2

    payload: dict[str, object] = {
        "ok": True,
        "output_path": str(output),
        "source_sha256_before": source_sha256_before,
        "service_date": snapshot.service_date.isoformat(),
        "snapshot_id": snapshot.snapshot_id,
        "event_id": event.event_id,
        "final_sessions": snapshot.final_session_count,
        "replacements": snapshot.replacement_count,
        "treatment_suspensions": snapshot.treatment_suspension_count,
        "unresolved_sessions": snapshot.unresolved_sessions,
        "therapist_history_rows": len(therapist_rows),
        "application_contract_source_components": source_components,
        "application_contract_output_components": len(output_contract.component_names),
        "source_unchanged": True,
    }
    _write_response(args.response, payload)

    print("CLOSE DAY PREVIEW VERIFIED")
    print(f"Date: {snapshot.service_date.isoformat()}")
    print(f"Snapshot: {snapshot.snapshot_id}")
    print(f"Event: {event.event_id}")
    print(f"Final sessions: {snapshot.final_session_count}")
    print(f"Therapist/day history rows: {len(therapist_rows)}")
    print(f"Unresolved sessions: {snapshot.unresolved_sessions}")
    print("Input unchanged: True")
    print(
        "Application contract: "
        f"{source_components} -> {len(output_contract.component_names)} components"
    )
    print(f"Output: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
