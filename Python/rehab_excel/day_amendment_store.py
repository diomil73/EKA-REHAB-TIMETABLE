from __future__ import annotations

from pathlib import Path
import sys
from typing import Iterable

from rehab_core.day_amendment import DayAmendmentPlan
from rehab_core.day_finalization import FinalizedSnapshotStatus
from rehab_core.therapist_day_history import TherapistDayHistoryRow
from rehab_excel.day_history_event_store import (
    HEADERS as EVENT_HEADERS,
    SHEET_NAME as EVENT_SHEET,
    _event_from_values,
    _values_for_com as event_values_for_com,
)
from rehab_excel.finalized_day_snapshots import (
    HEADERS as SNAPSHOT_HEADERS,
    SHEET_NAME as SNAPSHOT_SHEET,
    _semantic_key,
    _snapshot_from_values,
    _values_for_com as snapshot_values_for_com,
)
from rehab_excel.therapist_day_history_store import (
    HEADERS as THERAPIST_HEADERS,
    SHEET_NAME as THERAPIST_SHEET,
    _row_from_values,
    _values_for_com as therapist_values_for_com,
)


class DayAmendmentStoreError(RuntimeError):
    """Raised when a closed-day amendment cannot be persisted atomically."""


def _validate_plan_for_persistence(
    plan: DayAmendmentPlan,
    therapist_rows: Iterable[TherapistDayHistoryRow],
) -> tuple[TherapistDayHistoryRow, ...]:
    rows = tuple(therapist_rows)
    previous = plan.previous_snapshot
    superseded = plan.superseded_snapshot
    replacement = plan.replacement_snapshot

    if previous.status != FinalizedSnapshotStatus.CURRENT_FINAL:
        raise DayAmendmentStoreError("Previous snapshot must be CURRENT_FINAL")
    if superseded.snapshot_id != previous.snapshot_id:
        raise DayAmendmentStoreError("Superseded snapshot must preserve the previous SnapshotId")
    if superseded.status != FinalizedSnapshotStatus.SUPERSEDED:
        raise DayAmendmentStoreError("Superseded snapshot must have SUPERSEDED status")
    if replacement.status != FinalizedSnapshotStatus.CURRENT_FINAL:
        raise DayAmendmentStoreError("Replacement snapshot must be CURRENT_FINAL")
    if replacement.service_date != previous.service_date:
        raise DayAmendmentStoreError("Replacement snapshot must use the same service date")
    if replacement.version != previous.version + 1:
        raise DayAmendmentStoreError("Replacement snapshot version must increment by exactly one")
    if replacement.supersedes_snapshot_id != previous.snapshot_id:
        raise DayAmendmentStoreError("Replacement snapshot must supersede the previous SnapshotId")
    if replacement.unresolved_sessions != 0:
        raise DayAmendmentStoreError("Replacement snapshot cannot contain unresolved sessions")

    if plan.started_event.service_date != previous.service_date:
        raise DayAmendmentStoreError("Amendment-started event has the wrong service date")
    if plan.started_event.related_snapshot_id != previous.snapshot_id:
        raise DayAmendmentStoreError("Amendment-started event must reference the previous snapshot")
    if plan.finalized_event.service_date != previous.service_date:
        raise DayAmendmentStoreError("Amendment-finalized event has the wrong service date")
    if plan.finalized_event.related_event_id != plan.started_event.event_id:
        raise DayAmendmentStoreError("Amendment-finalized event must reference the started event")
    if plan.finalized_event.related_snapshot_id != replacement.snapshot_id:
        raise DayAmendmentStoreError("Amendment-finalized event must reference the replacement snapshot")

    keys = {(row.snapshot_id, row.therapist_id) for row in rows}
    if len(keys) != len(rows):
        raise DayAmendmentStoreError("Duplicate therapist rows in amendment batch")
    if any(row.snapshot_id != replacement.snapshot_id for row in rows):
        raise DayAmendmentStoreError("All therapist rows must belong to the replacement snapshot")
    if any(row.service_date != replacement.service_date for row in rows):
        raise DayAmendmentStoreError("All therapist rows must use the replacement service date")
    return rows


def _find_sheet(workbook, name: str):
    for sheet in workbook.Worksheets:
        if str(sheet.Name).casefold() == name.casefold():
            return sheet
    raise DayAmendmentStoreError(f"Required history sheet is missing: {name}")


def _require_headers(ws, expected: tuple[str, ...], *, sheet_name: str) -> None:
    actual = tuple(
        str(ws.Cells(1, col).Value or "").strip()
        for col in range(1, len(expected) + 1)
    )
    if actual != expected:
        raise DayAmendmentStoreError(
            f"{sheet_name} headers do not match authoritative schema: {actual!r}"
        )


def _last_data_row(ws) -> int:
    return max(int(ws.UsedRange.Rows.Count), 1)


def persist_day_amendment(
    workbook_path: str | Path,
    plan: DayAmendmentPlan,
    therapist_rows: Iterable[TherapistDayHistoryRow],
) -> dict[str, object]:
    """Persist one closed-day amendment in one native-Excel transaction.

    Nothing is saved until every snapshot, therapist-history and event row has
    passed semantic read-back. Any exception closes the workbook without saving,
    so the previous CURRENT_FINAL state remains authoritative on disk.
    """

    rows = _validate_plan_for_persistence(plan, therapist_rows)
    if sys.platform != "win32":
        raise DayAmendmentStoreError(
            "Closed-day amendment persistence requires Windows + Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DayAmendmentStoreError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise DayAmendmentStoreError(f"Workbook not found: {path}")

    excel = None
    workbook = None
    saved = False
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)

        snapshot_ws = _find_sheet(workbook, SNAPSHOT_SHEET)
        therapist_ws = _find_sheet(workbook, THERAPIST_SHEET)
        event_ws = _find_sheet(workbook, EVENT_SHEET)
        for ws in (snapshot_ws, therapist_ws, event_ws):
            try:
                ws.Unprotect()
            except Exception:
                pass

        _require_headers(snapshot_ws, SNAPSHOT_HEADERS, sheet_name=SNAPSHOT_SHEET)
        _require_headers(therapist_ws, THERAPIST_HEADERS, sheet_name=THERAPIST_SHEET)
        _require_headers(event_ws, EVENT_HEADERS, sheet_name=EVENT_SHEET)

        previous_row = None
        existing_snapshot_ids: set[str] = set()
        current_for_date: list[str] = []
        snapshot_last = _last_data_row(snapshot_ws)
        for row_number in range(2, snapshot_last + 1):
            values = tuple(
                snapshot_ws.Cells(row_number, col).Value
                for col in range(1, len(SNAPSHOT_HEADERS) + 1)
            )
            if not any(value is not None and str(value).strip() for value in values):
                continue
            existing = _snapshot_from_values(values, row=row_number)
            existing_snapshot_ids.add(existing.snapshot_id)
            if existing.snapshot_id == plan.previous_snapshot.snapshot_id:
                if _semantic_key(existing) != _semantic_key(plan.previous_snapshot):
                    raise DayAmendmentStoreError(
                        "Persisted previous snapshot no longer matches the amendment plan"
                    )
                previous_row = row_number
            if (
                existing.service_date == plan.previous_snapshot.service_date
                and existing.status == FinalizedSnapshotStatus.CURRENT_FINAL
            ):
                current_for_date.append(existing.snapshot_id)

        if previous_row is None:
            raise DayAmendmentStoreError("Previous snapshot was not found in workbook history")
        if current_for_date != [plan.previous_snapshot.snapshot_id]:
            raise DayAmendmentStoreError(
                "Expected the previous snapshot to be the only CURRENT_FINAL for the service date"
            )
        if plan.replacement_snapshot.snapshot_id in existing_snapshot_ids:
            raise DayAmendmentStoreError("Replacement SnapshotId already exists")

        therapist_last = _last_data_row(therapist_ws)
        existing_therapist_keys: set[tuple[str, str]] = set()
        for row_number in range(2, therapist_last + 1):
            values = tuple(
                therapist_ws.Cells(row_number, col).Value
                for col in range(1, len(THERAPIST_HEADERS) + 1)
            )
            if not any(value is not None and str(value).strip() for value in values):
                continue
            existing = _row_from_values(values, row=row_number)
            existing_therapist_keys.add((existing.snapshot_id, existing.therapist_id))
        requested_therapist_keys = {(row.snapshot_id, row.therapist_id) for row in rows}
        overlap = requested_therapist_keys & existing_therapist_keys
        if overlap:
            raise DayAmendmentStoreError(
                f"Replacement therapist/day rows already exist: {sorted(overlap)!r}"
            )

        event_last = _last_data_row(event_ws)
        existing_event_ids: set[str] = set()
        for row_number in range(2, event_last + 1):
            values = tuple(
                event_ws.Cells(row_number, col).Value
                for col in range(1, len(EVENT_HEADERS) + 1)
            )
            if not any(value is not None and str(value).strip() for value in values):
                continue
            existing_event_ids.add(_event_from_values(values, row=row_number).event_id)
        requested_event_ids = {
            plan.started_event.event_id,
            plan.finalized_event.event_id,
        }
        if requested_event_ids & existing_event_ids:
            raise DayAmendmentStoreError("One or more amendment EventIds already exist")

        # Mutation begins only after every preflight has passed.
        snapshot_ws.Cells(previous_row, 4).Value2 = FinalizedSnapshotStatus.SUPERSEDED.value
        previous_readback = tuple(
            snapshot_ws.Cells(previous_row, col).Value
            for col in range(1, len(SNAPSHOT_HEADERS) + 1)
        )
        persisted_superseded = _snapshot_from_values(previous_readback, row=previous_row)
        if _semantic_key(persisted_superseded) != _semantic_key(plan.superseded_snapshot):
            raise DayAmendmentStoreError("Superseded snapshot read-back mismatch")

        replacement_row = max(snapshot_last + 1, 2)
        for col, value in enumerate(snapshot_values_for_com(plan.replacement_snapshot), start=1):
            snapshot_ws.Cells(replacement_row, col).Value2 = value
        replacement_readback = tuple(
            snapshot_ws.Cells(replacement_row, col).Value
            for col in range(1, len(SNAPSHOT_HEADERS) + 1)
        )
        persisted_replacement = _snapshot_from_values(replacement_readback, row=replacement_row)
        if _semantic_key(persisted_replacement) != _semantic_key(plan.replacement_snapshot):
            raise DayAmendmentStoreError("Replacement snapshot read-back mismatch")

        therapist_start = max(therapist_last + 1, 2)
        for offset, row in enumerate(rows):
            target_row = therapist_start + offset
            for col, value in enumerate(therapist_values_for_com(row), start=1):
                therapist_ws.Cells(target_row, col).Value2 = value
            readback = tuple(
                therapist_ws.Cells(target_row, col).Value
                for col in range(1, len(THERAPIST_HEADERS) + 1)
            )
            if _row_from_values(readback, row=target_row) != row:
                raise DayAmendmentStoreError(
                    f"Therapist/day amendment read-back mismatch for {row.therapist_id!r}"
                )

        event_start = max(event_last + 1, 2)
        events = (plan.started_event, plan.finalized_event)
        for offset, event in enumerate(events):
            target_row = event_start + offset
            for col, value in enumerate(event_values_for_com(event), start=1):
                event_ws.Cells(target_row, col).Value2 = value
            readback = tuple(
                event_ws.Cells(target_row, col).Value
                for col in range(1, len(EVENT_HEADERS) + 1)
            )
            if _event_from_values(readback, row=target_row) != event:
                raise DayAmendmentStoreError(
                    f"Day-history amendment read-back mismatch for {event.event_id!r}"
                )

        for ws in (snapshot_ws, therapist_ws, event_ws):
            ws.Visible = 2
            ws.Protect()

        workbook.Save()
        saved = True
        return {
            "ok": True,
            "service_date": plan.replacement_snapshot.service_date.isoformat(),
            "previous_snapshot_id": plan.previous_snapshot.snapshot_id,
            "replacement_snapshot_id": plan.replacement_snapshot.snapshot_id,
            "replacement_version": plan.replacement_snapshot.version,
            "therapist_history_rows": len(rows),
            "events_written": 2,
        }
    except DayAmendmentStoreError:
        raise
    except Exception as exc:
        raise DayAmendmentStoreError(f"Could not persist closed-day amendment: {exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False if not saved else False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
