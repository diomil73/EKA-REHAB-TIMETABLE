from __future__ import annotations

"""Version-linked, evidence-first therapist working-day attendance contract.

This domain module has no Excel persistence and does not infer an absence from
missing sessions or from gaps in finalized operational snapshots. Callers must
supply the explicit expected-working-day calendar and independently verified
attendance facts before a ratio can be reported.
"""

from dataclasses import dataclass
from datetime import date
from enum import Enum
from typing import Iterable


class AttendanceHistoryError(ValueError):
    """Incomplete, contradictory or unproven attendance evidence."""


class AttendanceOutcome(str, Enum):
    PRESENT = "PRESENT"
    ABSENT = "ABSENT"
    NONWORKING = "NONWORKING"


@dataclass(frozen=True)
class TherapistAttendanceFact:
    service_date: date
    therapist_id: str
    snapshot_id: str
    outcome: AttendanceOutcome
    source: str
    reason_code: str | None = None

    def __post_init__(self) -> None:
        if not self.therapist_id.strip():
            raise AttendanceHistoryError("Therapist ID must be explicit")
        if not self.snapshot_id.strip():
            raise AttendanceHistoryError("CURRENT_FINAL snapshot ID required")
        if not self.source.strip():
            raise AttendanceHistoryError("Verified attendance provenance required")
        if self.outcome == AttendanceOutcome.ABSENT and not (
            self.reason_code and self.reason_code.strip()
        ):
            raise AttendanceHistoryError("Absence reason must be recorded")


@dataclass(frozen=True)
class TherapistAttendanceSummary:
    therapist_id: str
    working_days: int
    present_days: int
    absent_days: int

    @property
    def presence_ratio(self) -> float | None:
        return self.present_days / self.working_days if self.working_days else None


def summarize_attendance(
    therapist_id: str,
    expected_working_dates: Iterable[date],
    facts: Iterable[TherapistAttendanceFact],
    *,
    current_final_snapshot_ids: dict[date, str],
) -> TherapistAttendanceSummary:
    """Fail closed unless every expected working day has one proven current fact.

    The *caller* supplies verified calendar dates excluding weekends, holidays,
    and approved non-working periods. No synthetic weekdays or zero absences.
    """

    therapist = therapist_id.strip()
    if not therapist:
        raise AttendanceHistoryError("Therapist ID required")
    days = tuple(expected_working_dates)
    if len(set(days)) != len(days):
        raise AttendanceHistoryError("Duplicate expected working date")
    expected = set(days)
    by_date: dict[date, TherapistAttendanceFact] = {}
    for fact in facts:
        if fact.therapist_id.casefold() != therapist.casefold():
            raise AttendanceHistoryError("Mixed therapist evidence")
        if fact.service_date in by_date:
            raise AttendanceHistoryError("Duplicate attendance fact for date")
        if fact.service_date not in expected:
            raise AttendanceHistoryError("Attendance fact lies outside expected working dates")
        if fact.outcome == AttendanceOutcome.NONWORKING:
            raise AttendanceHistoryError("Nonworking day included in working denominator")
        if current_final_snapshot_ids.get(fact.service_date) != fact.snapshot_id:
            raise AttendanceHistoryError("Attendance fact is not linked to CURRENT_FINAL snapshot")
        by_date[fact.service_date] = fact
    missing = expected - set(by_date)
    if missing:
        raise AttendanceHistoryError(
            "Missing verified attendance facts: "
            + ", ".join(sorted(day.isoformat() for day in missing))
        )
    present = sum(f.outcome == AttendanceOutcome.PRESENT for f in by_date.values())
    absent = sum(f.outcome == AttendanceOutcome.ABSENT for f in by_date.values())
    if present + absent != len(expected):
        raise AttendanceHistoryError("Working-day coverage mismatch")
    return TherapistAttendanceSummary(therapist, len(expected), present, absent)
