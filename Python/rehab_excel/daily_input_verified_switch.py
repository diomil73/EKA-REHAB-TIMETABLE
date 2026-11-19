"""Join both daily-input archives before a native date switch is permitted.

No scheduling approval, workbook mutation or UI action. The caller must obtain
all inputs from one unchanged authoritative workbook and rebuild projections
from incoming.applied_result (never incoming.draft). Replacement validation,
projection readback and native rollback remain additional transaction gates.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .daily_input_applied_archive import APPLIED_SHEET, applied_archive_rows, read_applied_archive
from .daily_input_applied_state import VerifiedDayInput, verify_day_input
from .daily_input_day_switch import DailyInputDaySwitchError, DailyInputDaySwitchPlan
from .daily_input_draft_archive import ARCHIVE_SHEET, archive_rows, capture_daily_input, read_archive_sheet
from .daily_input_reader import _parse_date


@dataclass(frozen=True)
class VerifiedSwitchInputs:
    outgoing: VerifiedDayInput = field(repr=False)
    incoming: VerifiedDayInput = field(repr=False)
    draft_rows: tuple = field(repr=False)
    applied_rows: tuple = field(repr=False)


def verify_switch_inputs(*, plan, applied_grids, ledger_sheet, patients, sessions,
                         therapist_names, replacement_record_dates,
                         initialize_new_target=False):
    """Check both dates and preserve exact draft and applied membership.

    A never-used target can be initialized explicitly, with no existing draft,
    applied record, ledger row or replacement record (including restored).
    Missing evidence for a previously used day is not an empty programme.
    Finalized-date and source-hash checks belong to plan preparation.
    """
    if not isinstance(plan, DailyInputDaySwitchPlan):
        raise DailyInputDaySwitchError("Expected validated date-switch plan")
    existing = tuple(applied_grids)
    applied_archive_rows(existing)
    by_date = {grid.service_date: grid for grid in existing}
    source, target = plan.outgoing.service_date, plan.incoming.service_date
    if source == target or source not in by_date:
        raise DailyInputDaySwitchError("Missing outgoing applied evidence")
    had_draft = any(d.service_date == target for d in plan.saved_before)
    dates = tuple(replacement_record_dates)
    if any(type(day) is not date for day in dates):
        raise DailyInputDaySwitchError("Invalid replacement date inventory")
    if target not in by_date:
        if (initialize_new_target is not True or had_draft or target in dates
                or archive_rows((plan.incoming,)) != archive_rows((plan.incoming.empty_for(target),))):
            raise DailyInputDaySwitchError("Missing incoming applied evidence")
        for row in ledger_sheet.iter_rows(min_row=2, max_col=12, values_only=True):
            if any(v is not None for v in row) and _parse_date(row[1]) == target:
                raise DailyInputDaySwitchError("Target already has daily change history")
        by_date[target] = plan.incoming.empty_for(target)
    elif initialize_new_target or not had_draft:
        raise DailyInputDaySwitchError("Existing target requires its stored draft and applied evidence")

    patients, sessions, therapist_names = tuple(patients), tuple(sessions), tuple(therapist_names)
    def verify(grid):
        return verify_day_input(draft=grid, applied_grid=by_date[grid.service_date],
                                ledger_sheet=ledger_sheet, patients=patients,
                                sessions=(s for s in sessions if s.session_date == grid.service_date),
                                therapist_names=therapist_names)
    outgoing = verify(plan.outgoing)
    incoming = verify(plan.incoming)
    drafts = {grid.service_date: grid for grid in plan.saved_after}
    drafts[target] = plan.incoming
    return VerifiedSwitchInputs(outgoing, incoming, archive_rows(drafts.values()),
                                applied_archive_rows(by_date.values()))


def verify_switch_storage_readback(workbook, expected: VerifiedSwitchInputs):
    """Post-save check on both complete archives and the visible input grid."""
    if any(name not in workbook.sheetnames for name in ("DAILY_INPUT", ARCHIVE_SHEET, APPLIED_SHEET)):
        raise DailyInputDaySwitchError("Date-switch storage missing after save")
    raw = workbook[ARCHIVE_SHEET]
    if raw.sheet_state != "veryHidden" or not raw.protection.sheet:
        raise DailyInputDaySwitchError("Draft archive is not protected and veryHidden")
    visible = capture_daily_input(workbook["DAILY_INPUT"], parse_date=_parse_date)
    if (archive_rows((visible,)) != archive_rows((expected.incoming.draft,))
            or archive_rows(read_archive_sheet(raw)) != expected.draft_rows
            or applied_archive_rows(read_applied_archive(workbook[APPLIED_SHEET])) != expected.applied_rows):
        raise DailyInputDaySwitchError("Date-switch storage readback mismatch")
