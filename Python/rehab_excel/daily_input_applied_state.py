"""Read-only boundary between draft input and last applied input.

The caller must obtain applied_grid from a verified successful apply/restore
transaction, not from a draft, the rendered programme, or legacy row guessing.
The existing ledger corroborates membership and identities, but does not
contain enough data to reconstruct the original grid or prove all cell values.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from .authoritative_commit import file_sha256
from .application_contract import DAILY_CHANGE_LEDGER_HEADERS
from .daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError
from .daily_input_draft_archive import archive_rows
from .daily_input_reader import (
    DailyInputReadResult, DailyInputSourceChange, _parse_date,
    interpret_daily_input_grid, read_daily_input,
)


def active_input_ledger_keys(ws, service_date: date) -> tuple:
    """Read active source changes for exactly one date; never mutate display rows."""
    if type(service_date) is not date:
        raise DailyInputDaySwitchError("Explicit service date required")
    if ws.title != "DAILY_CHANGE_LEDGER":
        raise DailyInputDaySwitchError("Expected canonical daily change ledger")
    width = len(DAILY_CHANGE_LEDGER_HEADERS)
    if tuple(ws.cell(1, c).value for c in range(1, width + 1)) != DAILY_CHANGE_LEDGER_HEADERS:
        raise DailyInputDaySwitchError("Invalid daily change ledger schema")
    keys = {}
    for cells in ws.iter_rows(min_row=2, max_col=width):
        values = tuple(cell.value for cell in cells)
        if all(value is None for value in values):
            continue
        if any(cell.data_type in {"f", "e"} for cell in cells[:11]):
            raise DailyInputDaySwitchError("Invalid daily change ledger cell")
        change_id, raw_date, kind, subject, display, _, _, reason, row, status, _, _ = values
        if _parse_date(raw_date) != service_date:
            continue
        if status == "RESTORED":
            continue
        if status != "ACTIVE":
            raise DailyInputDaySwitchError("Unknown daily change status")
        if kind not in {"therapist_absence", "patient_absence", "session_cancellation"}:
            raise DailyInputDaySwitchError("Unsupported active daily input change type")
        if any(type(v) is not str or not v.strip() for v in (change_id, subject, display, reason)):
            raise DailyInputDaySwitchError("Incomplete active daily input change")
        if type(row) is not int or row < 7 or change_id in keys:
            raise DailyInputDaySwitchError("Invalid or duplicate active daily input change")
        keys[change_id] = (change_id, kind, subject, display, reason, row)
    return tuple(keys[key] for key in sorted(keys))


@dataclass(frozen=True)
class VerifiedDayInput:
    draft: DailyInputDayDraft = field(repr=False)
    applied_grid: DailyInputDayDraft = field(repr=False)
    applied_result: DailyInputReadResult = field(repr=False)

    @property
    def has_unapplied_edits(self) -> bool:
        # Typed serialization distinguishes even equal Python values of
        # different types (True/1, 1/1.0); this is a UI dirty flag, not approval.
        return archive_rows((self.draft,)) != archive_rows((self.applied_grid,))


def verify_day_input(
    *, draft: DailyInputDayDraft, applied_grid: DailyInputDayDraft | None,
    ledger_sheet, patients, sessions, therapist_names=(),
) -> VerifiedDayInput:
    """Reinterpret only verified applied input, preserving the pending draft.

    Missing applied evidence fails closed, even on a clean-looking day.
    An explicitly initialized empty applied grid is allowed when the active
    ledger has no source changes. Replacement facts are separate existing
    date-keyed records and must also be checked by the projection caller.
    """
    if not isinstance(draft, DailyInputDayDraft):
        raise DailyInputDaySwitchError("Expected literal draft")
    if not isinstance(applied_grid, DailyInputDayDraft):
        raise DailyInputDaySwitchError("Applied input evidence is required")
    if applied_grid.service_date != draft.service_date or applied_grid.layout != draft.layout:
        raise DailyInputDaySwitchError("Applied input date or layout mismatch")
    active = active_input_ledger_keys(ledger_sheet, draft.service_date)
    result = interpret_daily_input_grid(
        applied_grid, patients=patients, sessions=sessions,
        therapist_names=therapist_names, allow_empty=True,
    )
    expected = tuple(sorted(
        (c.change_id, c.change_type, c.subject_id, c.display_name, c.reason, c.source_row)
        for c in result.source_changes
    ))
    if expected != active:
        raise DailyInputDaySwitchError("Applied grid does not match active daily input ledger")
    return VerifiedDayInput(draft, applied_grid, result)

@dataclass(frozen=True)
class InputLedgerSavedReadback:
    """Describes a verified OOXML readback, never an Apply/phase authorization."""

    service_date: date
    workbook_sha256: str
    source_change_count: int
    active_ledger_count: int


def _source_change_keys(changes):
    if not all(type(change) is DailyInputSourceChange for change in changes):
        raise DailyInputDaySwitchError("Invalid validated DAILY_INPUT source changes")
    keys = tuple(sorted((
        change.change_id, change.change_type, change.subject_id,
        change.display_name, change.reason, change.source_row,
    ) for change in changes))
    if len(set(keys)) != len(keys) or len({key[0] for key in keys}) != len(keys):
        raise DailyInputDaySwitchError("Duplicate source changes in DAILY_INPUT readback")
    return keys


def verify_input_ledger_saved_readback(
    workbook_path: str | Path,
    *,
    service_date: date,
    validated_changes: tuple[DailyInputSourceChange, ...],
    patients,
    sessions,
    therapist_names=(),
) -> InputLedgerSavedReadback:
    """Compare saved DAILY_INPUT and DAILY_CHANGE_LEDGER against preflight facts.

    This is a strict read-only saved OOXML consistency check for the proposed
    INPUT_AND_LEDGER_COMMITTED phase. The expected changes MUST originate from
    the earlier trusted preflight, not be reconstructed from the same reopened
    ledger. Source/ledger identities are fully compared, not merely counted.
    It is NOT a native Excel reopen, trusted receipt, finalization, operation
    journal append, or permission to lift the Apply gate.
    """
    if type(service_date) is not date:
        raise DailyInputDaySwitchError("Explicit DAILY_INPUT service date required")
    if type(validated_changes) is not tuple or not validated_changes:
        raise DailyInputDaySwitchError("Non-empty validated preflight changes required")
    expected = _source_change_keys(validated_changes)
    source = Path(workbook_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise DailyInputDaySwitchError("Existing XLSM workbook required")
    source_sha = file_sha256(source)

    try:
        actual = read_daily_input(
            source, patients=tuple(patients), sessions=tuple(sessions),
            therapist_names=tuple(therapist_names),
        )
        if actual.target_date != service_date:
            raise DailyInputDaySwitchError("Saved DAILY_INPUT date changed")
        if _source_change_keys(actual.source_changes) != expected:
            raise DailyInputDaySwitchError("Saved DAILY_INPUT differs from preflight source changes")

        wb = load_workbook(source, read_only=False, data_only=False, keep_vba=True)
        try:
            if "DAILY_CHANGE_LEDGER" not in wb.sheetnames:
                raise DailyInputDaySwitchError("Saved daily change ledger missing")
            ledger = wb["DAILY_CHANGE_LEDGER"]
            if ledger.sheet_state != "veryHidden":
                raise DailyInputDaySwitchError("Saved daily change ledger not veryHidden")
            active = active_input_ledger_keys(ledger, service_date)
        finally:
            wb.close()
        if active != expected:
            raise DailyInputDaySwitchError("Saved active ledger differs from validated DAILY_INPUT")
    except DailyInputDaySwitchError:
        raise
    except Exception as exc:
        raise DailyInputDaySwitchError("Saved input and ledger readback invalid") from exc
    if file_sha256(source) != source_sha:
        raise DailyInputDaySwitchError("Workbook changed during input and ledger readback")
    return InputLedgerSavedReadback(
        service_date=service_date,
        workbook_sha256=source_sha,
        source_change_count=len(expected),
        active_ledger_count=len(active),
    )

