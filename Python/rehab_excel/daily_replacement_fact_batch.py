from __future__ import annotations

from pathlib import Path
import sys
from typing import Iterable

from .daily_replacement_facts import (
    ACTIVE_STATUS,
    HEADERS,
    RESTORED_STATUS,
    SHEET_NAME,
    DailyReplacementFactError,
)


def validate_batch_restore_request(
    fact_rows: Iterable[object],
    replacement_ids: Iterable[str],
) -> tuple[str, ...]:
    """Validate an entire restore set before any workbook mutation occurs."""

    wanted = tuple(str(value).strip() for value in replacement_ids)
    if not wanted:
        raise DailyReplacementFactError("At least one ReplacementId is required")
    if any(not value for value in wanted):
        raise DailyReplacementFactError("ReplacementId cannot be blank")
    if len(set(wanted)) != len(wanted):
        raise DailyReplacementFactError("Duplicate ReplacementId in restore request")

    by_id: dict[str, object] = {}
    for fact in fact_rows:
        replacement_id = str(getattr(fact, "replacement_id", "") or "").strip()
        if not replacement_id:
            continue
        if replacement_id in by_id:
            raise DailyReplacementFactError(
                f"Duplicate ReplacementId {replacement_id!r} in persisted facts"
            )
        by_id[replacement_id] = fact

    for replacement_id in wanted:
        fact = by_id.get(replacement_id)
        if fact is None:
            raise DailyReplacementFactError(
                f"ReplacementId {replacement_id!r} not found"
            )
        status = str(getattr(fact, "status", ACTIVE_STATUS) or ACTIVE_STATUS).strip().upper()
        if status == RESTORED_STATUS:
            raise DailyReplacementFactError(
                f"ReplacementId {replacement_id!r} is already restored"
            )
        if status != ACTIVE_STATUS:
            raise DailyReplacementFactError(
                f"ReplacementId {replacement_id!r} has unsupported status {status!r}"
            )

    return wanted


def restore_daily_replacement_facts(
    workbook_path: str | Path,
    replacement_ids: Iterable[str],
) -> tuple[str, ...]:
    """Atomically mark a validated set of ACTIVE facts as RESTORED.

    Every requested id is resolved and validated first. Only after the entire
    set is proven safe are any status cells changed, followed by one workbook
    save. A stale/missing member therefore cannot produce a partial restore.
    """

    if sys.platform != "win32":
        raise DailyReplacementFactError(
            "Daily replacement batch restore requires Windows + Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyReplacementFactError("pywin32 is required") from exc

    requested = tuple(str(value).strip() for value in replacement_ids)
    path = Path(workbook_path).resolve()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        try:
            ws = workbook.Worksheets(SHEET_NAME)
        except Exception as exc:
            raise DailyReplacementFactError(
                f"Workbook has no {SHEET_NAME} sheet"
            ) from exc

        actual_headers = tuple(
            str(ws.Cells(1, col).Value or "").strip()
            for col in range(1, len(HEADERS) + 1)
        )
        if actual_headers != HEADERS:
            raise DailyReplacementFactError(
                f"{SHEET_NAME} headers do not match authoritative schema: {actual_headers!r}"
            )

        last_row = max(int(ws.UsedRange.Rows.Count), 1)
        row_by_id: dict[str, int] = {}
        status_by_id: dict[str, str] = {}
        fact_views: list[object] = []

        class _FactView:
            def __init__(self, replacement_id: str, status: str) -> None:
                self.replacement_id = replacement_id
                self.status = status

        for row_number in range(2, last_row + 1):
            replacement_id = str(ws.Cells(row_number, 1).Value or "").strip()
            if not replacement_id:
                continue
            if replacement_id in row_by_id:
                raise DailyReplacementFactError(
                    f"Duplicate ReplacementId {replacement_id!r} in persisted facts"
                )
            status = str(ws.Cells(row_number, 10).Value or ACTIVE_STATUS).strip().upper()
            row_by_id[replacement_id] = row_number
            status_by_id[replacement_id] = status
            fact_views.append(_FactView(replacement_id, status))

        wanted = validate_batch_restore_request(fact_views, requested)

        # Mutation starts only after the whole request has passed validation.
        for replacement_id in wanted:
            ws.Cells(row_by_id[replacement_id], 10).Value = RESTORED_STATUS
        ws.Visible = 2  # xlSheetVeryHidden
        workbook.Save()
        return wanted
    except DailyReplacementFactError:
        raise
    except Exception as exc:
        raise DailyReplacementFactError(
            f"Could not batch-restore daily replacement facts: {exc}"
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
