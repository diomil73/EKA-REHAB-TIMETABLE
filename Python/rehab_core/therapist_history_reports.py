from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable

from .day_finalization import FinalizedDaySnapshot, FinalizedSnapshotStatus
from .therapist_day_history import TherapistDayHistoryRow


class TherapistHistoryReportError(RuntimeError):
    """Raised when historical therapist reporting cannot be derived safely."""


@dataclass(frozen=True)
class TherapistDailyReport:
    snapshot_id: str
    service_date: date
    snapshot_version: int
    therapist_id: str
    capacity_timeslots: int | None
    planned_occupied_timeslots: int
    planned_sessions: int
    final_occupied_timeslots: int
    final_active_sessions: int
    replacements_received: int
    released_by_cancellation_or_no_show: int
    treatment_suspensions_affecting_load: int
    approved_shared_or_over_capacity_events: int

    @property
    def planned_capacity_utilization(self) -> float | None:
        if self.capacity_timeslots is None or self.capacity_timeslots <= 0:
            return None
        return self.planned_occupied_timeslots / self.capacity_timeslots

    @property
    def actual_capacity_utilization(self) -> float | None:
        if self.capacity_timeslots is None or self.capacity_timeslots <= 0:
            return None
        return self.final_occupied_timeslots / self.capacity_timeslots

    @property
    def session_delta(self) -> int:
        return self.final_active_sessions - self.planned_sessions


def build_therapist_daily_reports(
    snapshots: Iterable[FinalizedDaySnapshot],
    therapist_rows: Iterable[TherapistDayHistoryRow],
    *,
    therapist_id: str | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
) -> tuple[TherapistDailyReport, ...]:
    """Build read-only daily therapist reports from finalized historical facts.

    Only CURRENT_FINAL snapshots are reportable. Superseded versions remain in
    history but are deliberately excluded so a v1 -> v2 amendment cannot double
    count the same service day.
    """

    if start_date is not None and end_date is not None and start_date > end_date:
        raise TherapistHistoryReportError("start_date cannot be after end_date")

    requested_therapist = therapist_id.strip() if therapist_id is not None else None
    if therapist_id is not None and not requested_therapist:
        raise TherapistHistoryReportError("therapist_id cannot be blank")

    current_by_date: dict[date, FinalizedDaySnapshot] = {}
    current_by_id: dict[str, FinalizedDaySnapshot] = {}
    for snapshot in snapshots:
        if snapshot.status != FinalizedSnapshotStatus.CURRENT_FINAL:
            continue
        if snapshot.unresolved_sessions != 0:
            raise TherapistHistoryReportError(
                f"Current finalized snapshot has unresolved sessions: {snapshot.snapshot_id}"
            )
        if snapshot.service_date in current_by_date:
            other = current_by_date[snapshot.service_date]
            raise TherapistHistoryReportError(
                "Multiple CURRENT_FINAL snapshots for service date "
                f"{snapshot.service_date.isoformat()}: "
                f"{other.snapshot_id}, {snapshot.snapshot_id}"
            )
        current_by_date[snapshot.service_date] = snapshot
        current_by_id[snapshot.snapshot_id] = snapshot

    reports: list[TherapistDailyReport] = []
    seen_pairs: set[tuple[str, str]] = set()
    for row in therapist_rows:
        snapshot = current_by_id.get(row.snapshot_id)
        if snapshot is None:
            continue
        if row.service_date != snapshot.service_date:
            raise TherapistHistoryReportError(
                f"Therapist row date does not match snapshot: {row.snapshot_id}"
            )
        if requested_therapist is not None and row.therapist_id != requested_therapist:
            continue
        if start_date is not None and row.service_date < start_date:
            continue
        if end_date is not None and row.service_date > end_date:
            continue

        pair = (row.snapshot_id, row.therapist_id)
        if pair in seen_pairs:
            raise TherapistHistoryReportError(
                f"Duplicate therapist/day history row: {row.snapshot_id} / {row.therapist_id}"
            )
        seen_pairs.add(pair)

        reports.append(
            TherapistDailyReport(
                snapshot_id=row.snapshot_id,
                service_date=row.service_date,
                snapshot_version=snapshot.version,
                therapist_id=row.therapist_id,
                capacity_timeslots=row.capacity_timeslots,
                planned_occupied_timeslots=row.planned_occupied_timeslots,
                planned_sessions=row.planned_sessions,
                final_occupied_timeslots=row.final_occupied_timeslots,
                final_active_sessions=row.final_active_sessions,
                replacements_received=row.replacements_received,
                released_by_cancellation_or_no_show=row.released_by_cancellation_or_no_show,
                treatment_suspensions_affecting_load=row.treatment_suspensions_affecting_load,
                approved_shared_or_over_capacity_events=row.approved_shared_or_over_capacity_events,
            )
        )

    reports.sort(key=lambda report: (report.service_date, report.therapist_id.casefold()))
    return tuple(reports)
