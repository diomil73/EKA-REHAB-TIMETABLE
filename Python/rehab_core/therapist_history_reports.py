from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
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


@dataclass(frozen=True)
class TherapistWeeklyReport:
    week_start: date
    week_end: date
    therapist_id: str
    reported_days: int
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


@dataclass(frozen=True)
class TherapistPeriodReport:
    period_start: date
    period_end: date
    therapist_id: str
    reported_days: int
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


def _capacity_total(rows: Iterable[TherapistDailyReport]) -> int | None:
    values = [row.capacity_timeslots for row in rows]
    if any(value is None for value in values):
        return None
    return sum(int(value) for value in values if value is not None)


def _period_report(
    rows: list[TherapistDailyReport],
    *,
    period_start: date,
    period_end: date,
    therapist_id: str,
) -> TherapistPeriodReport:
    dates = {row.service_date for row in rows}
    if len(dates) != len(rows):
        raise TherapistHistoryReportError(
            f"Duplicate daily report for therapist/date: {therapist_id} / {period_start.isoformat()}"
        )
    return TherapistPeriodReport(
        period_start=period_start,
        period_end=period_end,
        therapist_id=therapist_id,
        reported_days=len(rows),
        capacity_timeslots=_capacity_total(rows),
        planned_occupied_timeslots=sum(row.planned_occupied_timeslots for row in rows),
        planned_sessions=sum(row.planned_sessions for row in rows),
        final_occupied_timeslots=sum(row.final_occupied_timeslots for row in rows),
        final_active_sessions=sum(row.final_active_sessions for row in rows),
        replacements_received=sum(row.replacements_received for row in rows),
        released_by_cancellation_or_no_show=sum(
            row.released_by_cancellation_or_no_show for row in rows
        ),
        treatment_suspensions_affecting_load=sum(
            row.treatment_suspensions_affecting_load for row in rows
        ),
        approved_shared_or_over_capacity_events=sum(
            row.approved_shared_or_over_capacity_events for row in rows
        ),
    )


def build_therapist_weekly_reports(
    daily_reports: Iterable[TherapistDailyReport],
) -> tuple[TherapistWeeklyReport, ...]:
    """Aggregate authoritative daily reports into Monday-Sunday workload trends."""

    groups: dict[tuple[date, str], list[TherapistDailyReport]] = {}
    for report in daily_reports:
        week_start = report.service_date - timedelta(days=report.service_date.weekday())
        groups.setdefault((week_start, report.therapist_id), []).append(report)

    result: list[TherapistWeeklyReport] = []
    for (week_start, therapist_id), rows in groups.items():
        rows = sorted(rows, key=lambda item: item.service_date)
        period = _period_report(
            rows,
            period_start=week_start,
            period_end=week_start + timedelta(days=6),
            therapist_id=therapist_id,
        )
        result.append(
            TherapistWeeklyReport(
                week_start=period.period_start,
                week_end=period.period_end,
                therapist_id=period.therapist_id,
                reported_days=period.reported_days,
                capacity_timeslots=period.capacity_timeslots,
                planned_occupied_timeslots=period.planned_occupied_timeslots,
                planned_sessions=period.planned_sessions,
                final_occupied_timeslots=period.final_occupied_timeslots,
                final_active_sessions=period.final_active_sessions,
                replacements_received=period.replacements_received,
                released_by_cancellation_or_no_show=period.released_by_cancellation_or_no_show,
                treatment_suspensions_affecting_load=period.treatment_suspensions_affecting_load,
                approved_shared_or_over_capacity_events=period.approved_shared_or_over_capacity_events,
            )
        )

    result.sort(key=lambda report: (report.week_start, report.therapist_id.casefold()))
    return tuple(result)


def build_therapist_monthly_reports(
    daily_reports: Iterable[TherapistDailyReport],
) -> tuple[TherapistPeriodReport, ...]:
    """Aggregate authoritative daily reports by calendar month and therapist."""

    groups: dict[tuple[int, int, str], list[TherapistDailyReport]] = {}
    for report in daily_reports:
        key = (report.service_date.year, report.service_date.month, report.therapist_id)
        groups.setdefault(key, []).append(report)

    result: list[TherapistPeriodReport] = []
    for (year, month, therapist_id), rows in groups.items():
        rows = sorted(rows, key=lambda item: item.service_date)
        month_start = date(year, month, 1)
        next_month = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
        result.append(
            _period_report(
                rows,
                period_start=month_start,
                period_end=next_month - timedelta(days=1),
                therapist_id=therapist_id,
            )
        )

    result.sort(key=lambda report: (report.period_start, report.therapist_id.casefold()))
    return tuple(result)


def build_therapist_overall_reports(
    daily_reports: Iterable[TherapistDailyReport],
) -> tuple[TherapistPeriodReport, ...]:
    """Aggregate all available finalized history into one row per therapist."""

    groups: dict[str, list[TherapistDailyReport]] = {}
    for report in daily_reports:
        groups.setdefault(report.therapist_id, []).append(report)

    result: list[TherapistPeriodReport] = []
    for therapist_id, rows in groups.items():
        rows = sorted(rows, key=lambda item: item.service_date)
        result.append(
            _period_report(
                rows,
                period_start=rows[0].service_date,
                period_end=rows[-1].service_date,
                therapist_id=therapist_id,
            )
        )

    result.sort(key=lambda report: report.therapist_id.casefold())
    return tuple(result)
