"""Read-only preparation of an open-day transition in the same workbook.

The returned plan is private, bound to the exact source bytes. A separate
native transaction must validate that binding, preserve confirmed facts,
rebuild all operational projections, verify persisted rows and roll back on
failure. Never repurpose this plan as an applied scheduling decision.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from .authoritative_commit import file_sha256
from .daily_input_reader import _parse_date
from .daily_input_day_switch import (
    DailyInputDaySwitchError, DailyInputDaySwitchPlan, plan_open_day_switch,
)
from .daily_input_draft_archive import (
    ARCHIVE_SHEET, archive_rows, capture_daily_input, read_archive_sheet,
)
from .finalized_day_snapshots import read_finalized_day_snapshots


@dataclass(frozen=True)
class PreparedOpenDaySwitch:
    source_sha256: str = field(repr=False)
    day_plan: DailyInputDaySwitchPlan = field(repr=False)
    expected_archive_rows: tuple = field(repr=False)


def prepare_open_day_switch(
    workbook_path: str | Path, *, expected_current_date: date, target_date: date,
    initialize_archive: bool = False,
) -> PreparedOpenDaySwitch:
    """Read source only; existing history is never routed through draft editing."""
    path = Path(workbook_path).resolve()
    if not path.is_file() or path.suffix.lower() != ".xlsm":
        raise DailyInputDaySwitchError("Existing XLSM application source required")
    before = file_sha256(path)
    # Normal openpyxl reader exposes validation ranges. It is never saved.
    wb = load_workbook(path, read_only=False, data_only=False)
    try:
        if "DAILY_INPUT" not in wb.sheetnames:
            raise DailyInputDaySwitchError("Missing DAILY_INPUT")
        current = capture_daily_input(wb["DAILY_INPUT"], parse_date=_parse_date)
        if ARCHIVE_SHEET in wb.sheetnames:
            ws = wb[ARCHIVE_SHEET]
            if initialize_archive:
                raise DailyInputDaySwitchError("Archive initialization would overwrite existing data")
            if ws.sheet_state != "veryHidden" or not ws.protection.sheet:
                raise DailyInputDaySwitchError("Draft archive must be veryHidden and protected")
            saved = read_archive_sheet(ws)
        elif initialize_archive is True:
            # Explicit first-install authorization only. Never use as recovery
            # for a formerly existing, missing archive.
            saved = ()
        else:
            raise DailyInputDaySwitchError("Missing draft archive; explicit first installation required")
    finally:
        wb.close()
    # Use the existing authoritative history reader. Even a superseded-only
    # date is kept out of the open route; ambiguity is not a fresh draft.
    history = read_finalized_day_snapshots(path)
    plan = plan_open_day_switch(
        current=current, saved_drafts=saved,
        expected_current_date=expected_current_date, target_date=target_date,
        finalized_dates={snapshot.service_date for snapshot in history},
    )
    rows = archive_rows(plan.saved_after)
    if file_sha256(path) != before:
        raise DailyInputDaySwitchError("Source changed during date-switch preparation")
    return PreparedOpenDaySwitch(before, plan, rows)
