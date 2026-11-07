from __future__ import annotations

"""Strictly read-only, version-aware therapist attendance evidence reader.

This module does NOT create a sheet, write a workbook, infer holidays, or
silently replace missing attendance evidence with zero absences. A future
finalized-day writer must explicitly produce this schema.
"""

from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from rehab_core.therapist_attendance_history import (
    AttendanceHistoryError,
    AttendanceOutcome,
    TherapistAttendanceFact,
    TherapistAttendanceSummary,
    summarize_attendance,
)
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots


SHEET_NAME = "THERAPIST_ATTENDANCE_HISTORY"
HEADERS = (
    "SnapshotId",
    "ServiceDate",
    "TherapistId",
    "Outcome",
    "Source",
    "ReasonCode",
)


class TherapistAttendanceHistoryStoreError(RuntimeError):
    """Missing, incomplete, malformed or unversioned attendance evidence."""


def _source_date(value: object, *, row: int) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value or "").strip())
    except ValueError as exc:
        raise TherapistAttendanceHistoryStoreError(
            f"{SHEET_NAME} row {row}: invalid ServiceDate {value!r}"
        ) from exc


def read_therapist_attendance_history_facts(
    workbook_path: str | Path,
) -> tuple[TherapistAttendanceFact, ...]:
    """Require the explicit new historical source; never create it on read."""

    path = Path(workbook_path)
    wb = load_workbook(
        path,
        read_only=True,
        data_only=False,
        keep_vba=path.suffix.casefold() == ".xlsm",
    )
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise TherapistAttendanceHistoryStoreError(
                f"{SHEET_NAME} is missing; therapist absences remain unavailable"
            )
        ws = wb[SHEET_NAME]
        actual = tuple(
            str(ws.cell(1, col).value or "").strip()
            for col in range(1, len(HEADERS) + 1)
        )
        if actual != HEADERS:
            raise TherapistAttendanceHistoryStoreError(
                f"{SHEET_NAME}: schema header mismatch: {actual!r}"
            )
        result: list[TherapistAttendanceFact] = []
        seen: set[tuple[str, str, date]] = set()
        for row_num, values in enumerate(
            ws.iter_rows(min_row=2, min_col=1, max_col=len(HEADERS), values_only=True),
            start=2,
        ):
            if not any(v is not None and str(v).strip() for v in values):
                continue
            snapshot, when, therapist, outcome, source, reason = values
            try:
                status = AttendanceOutcome(str(outcome or "").strip().upper())
            except ValueError as exc:
                raise TherapistAttendanceHistoryStoreError(
                    f"{SHEET_NAME} row {row_num}: invalid Outcome {outcome!r}"
                ) from exc
            when_date = _source_date(when, row=row_num)
            try:
                fact = TherapistAttendanceFact(
                    service_date=when_date,
                    therapist_id=str(therapist or "").strip(),
                    snapshot_id=str(snapshot or "").strip(),
                    outcome=status,
                    source=str(source or "").strip(),
                    reason_code=str(reason).strip() if reason is not None and str(reason).strip() else None,
                )
            except AttendanceHistoryError as exc:
                raise TherapistAttendanceHistoryStoreError(
                    f"{SHEET_NAME} row {row_num}: {exc}"
                ) from exc
            key = (fact.snapshot_id.casefold(), fact.therapist_id.casefold(), fact.service_date)
            if key in seen:
                raise TherapistAttendanceHistoryStoreError(
                    f"{SHEET_NAME} row {row_num}: duplicate snapshot/therapist/date"
                )
            seen.add(key)
            result.append(fact)
        return tuple(result)
    finally:
        wb.close()


def read_verified_therapist_attendance(
    workbook_path: str | Path,
    *,
    therapist_id: str,
    expected_working_dates: Iterable[date],
) -> TherapistAttendanceSummary:
    """Count only facts for the latest CURRENT_FINAL version of each day.

    The expected-working-day calendar must come from an independently verified
    roster/holiday/leave policy. No default Mon-Fri guess is permitted.
    """

    expected = tuple(expected_working_dates)
    current = read_finalized_day_snapshots(workbook_path, current_only=True)
    current_ids = {snapshot.service_date: snapshot.snapshot_id for snapshot in current}
    if any(day not in current_ids for day in expected):
        raise TherapistAttendanceHistoryStoreError(
            "Expected working date without a CURRENT_FINAL operational snapshot"
        )
    all_facts = read_therapist_attendance_history_facts(workbook_path)
    selected = tuple(
        fact for fact in all_facts
        if fact.therapist_id.casefold() == therapist_id.strip().casefold()
        and fact.service_date in expected
        and fact.snapshot_id == current_ids.get(fact.service_date)
    )
    try:
        return summarize_attendance(
            therapist_id, expected, selected,
            current_final_snapshot_ids=current_ids,
        )
    except AttendanceHistoryError as exc:
        raise TherapistAttendanceHistoryStoreError(str(exc)) from exc
