"""Storage contract for the last successfully applied input of each open day.

This is not a scheduling authority. Use the existing native Apply/Restore
transaction and its readback as the only source of a newly applied grid.
These helpers perform no workbook writes and never turn draft-save into Apply.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from hashlib import sha256

from .daily_input_applied_state import verify_day_input
from .daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError
from .daily_input_draft_archive import archive_rows, decode_archive_rows

APPLIED_SHEET = "DAILY_INPUT_APPLIED"
HEADERS = ("ServiceDate", "Part", "Parts", "AppliedSHA256", "AppliedPayload")
_PREFIX = "applied-v1:"


def _digest(raw_digest: str) -> str:
    return sha256((_PREFIX + raw_digest).encode("ascii")).hexdigest()


def applied_archive_rows(grids):
    """Reuse the typed draft codec but keep the two storage roles distinct."""
    return tuple((day, part, total, _digest(digest), _PREFIX + digest + ":" + payload)
                 for day, part, total, digest, payload in archive_rows(grids))


def decode_applied_rows(rows):
    decoded = []
    for row in rows:
        if len(row) != 5:
            raise DailyInputDaySwitchError("Invalid applied archive row")
        day, part, total, digest, payload = row
        if type(payload) is not str or not payload.startswith(_PREFIX):
            raise DailyInputDaySwitchError("Applied archive role mismatch")
        raw_digest, sep, raw_payload = payload[len(_PREFIX):].partition(":")
        if (sep != ":" or len(raw_digest) != 64
                or any(c not in "0123456789abcdef" for c in raw_digest)
                or digest != _digest(raw_digest)):
            raise DailyInputDaySwitchError("Invalid applied archive digest")
        decoded.append((day, part, total, raw_digest, raw_payload))
    return decode_archive_rows(decoded)


def read_applied_archive(ws):
    if (ws.title != APPLIED_SHEET or ws.sheet_state != "veryHidden"
            or not ws.protection.sheet):
        raise DailyInputDaySwitchError("Expected protected, veryHidden applied archive")
    if tuple(ws.cell(1, col).value for col in range(1, 6)) != HEADERS:
        raise DailyInputDaySwitchError("Invalid applied archive header")
    rows = []
    for cells in ws.iter_rows(min_row=2, max_col=5):
        if any(c.data_type in {"f", "e"} for c in cells):
            raise DailyInputDaySwitchError("Formula or error in applied archive")
        row = tuple(c.value for c in cells)
        if any(v is not None for v in row):
            rows.append(row)
    return decode_applied_rows(rows)


@dataclass(frozen=True)
class AppliedArchiveUpdate:
    before: tuple[DailyInputDayDraft, ...] = field(repr=False)
    after: tuple[DailyInputDayDraft, ...] = field(repr=False)

    @property
    def expected_rows(self):
        return applied_archive_rows(self.after)


def plan_applied_archive_update(*, existing, expected_previous, candidate,
                                ledger_sheet, finalized_dates, patients,
                                sessions, therapist_names=()):
    """Plan an archive update within the Apply/Restore transaction, not on save.

    The candidate must be the exact validated grid used by that transaction;
    ledger_sheet must represent its post-apply state. The caller must not
    persist this plan until operational patches, replacement facts, and all
    relevant readbacks succeed; a failed transaction must roll back all of them.
    expected_previous=None is allowed only for an explicitly new record.
    """
    before = tuple(existing)
    applied_archive_rows(before)  # validate exact membership and duplicates
    if not isinstance(candidate, DailyInputDayDraft):
        raise DailyInputDaySwitchError("Applied candidate grid required")
    finals = tuple(finalized_dates)
    if any(type(day) is not date for day in finals) or candidate.service_date in finals:
        raise DailyInputDaySwitchError("Applied archive cannot update finalized day")
    by_day = {grid.service_date: grid for grid in before}
    prior = by_day.get(candidate.service_date)
    if expected_previous is not None and not isinstance(expected_previous, DailyInputDayDraft):
        raise DailyInputDaySwitchError("Invalid previous applied evidence")
    if applied_archive_rows(() if prior is None else (prior,)) != applied_archive_rows(
            () if expected_previous is None else (expected_previous,)):
        raise DailyInputDaySwitchError("Previous applied archive changed")
    if prior is not None and prior.layout != candidate.layout:
        raise DailyInputDaySwitchError("Applied grid layout changed")
    verify_day_input(draft=candidate, applied_grid=candidate, ledger_sheet=ledger_sheet,
                     patients=patients, sessions=sessions, therapist_names=therapist_names)
    by_day[candidate.service_date] = candidate
    return AppliedArchiveUpdate(before, tuple(by_day[day] for day in sorted(by_day)))


def verify_applied_archive_readback(ws, update: AppliedArchiveUpdate):
    """Check every date and typed cell, including records outside the target day."""
    if applied_archive_rows(read_applied_archive(ws)) != update.expected_rows:
        raise DailyInputDaySwitchError("Applied archive readback differs from expected state")
