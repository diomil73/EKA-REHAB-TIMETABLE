from datetime import date, datetime, timezone

import pytest

from rehab_core.day_finalization import FinalizedDaySnapshot, FinalizedSnapshotStatus
from rehab_core.therapist_day_history import TherapistDayHistoryRow
from rehab_core.therapist_history_reports import (
    TherapistHistoryReportError,
    build_therapist_daily_reports,
)


DAY = date(2026, 10, 30)


def _snapshot(
    snapshot_id: str,
    *,
    version: int,
    status: FinalizedSnapshotStatus,
    service_date: date = DAY,
    unresolved_sessions: int = 0,
) -> FinalizedDaySnapshot:
    return FinalizedDaySnapshot(
        snapshot_id=snapshot_id,
        service_date=service_date,
        version=version,
        status=status,
        finalized_at=datetime(2026, 11, 2, 21, 0, tzinfo=timezone.utc),
        supersedes_snapshot_id="snap-v1" if version > 1 else None,
        unresolved_sessions=unresolved_sessions,
        planned_session_count=6,
        final_session_count=5,
        patient_absence_count=1,
        cancellation_count=0,
        no_show_count=0,
        replacement_count=1,
        treatment_suspension_count=0,
        approved_exception_count=0,
    )


def _row(snapshot_id: str, therapist_id: str = "T1", *, service_date: date = DAY) -> TherapistDayHistoryRow:
    return TherapistDayHistoryRow(
        snapshot_id=snapshot_id,
        service_date=service_date,
        therapist_id=therapist_id,
        capacity_timeslots=6,
        planned_occupied_timeslots=5,
        planned_sessions=5,
        final_occupied_timeslots=4,
        final_active_sessions=4,
        replacements_received=1,
        released_by_cancellation_or_no_show=1,
        treatment_suspensions_affecting_load=0,
        approved_shared_or_over_capacity_events=0,
    )


def test_reports_use_only_current_final_version_after_amendment():
    snapshots = (
        _snapshot("snap-v1", version=1, status=FinalizedSnapshotStatus.SUPERSEDED),
        _snapshot("snap-v2", version=2, status=FinalizedSnapshotStatus.CURRENT_FINAL),
    )
    rows = (_row("snap-v1"), _row("snap-v2"))

    reports = build_therapist_daily_reports(snapshots, rows)

    assert len(reports) == 1
    assert reports[0].snapshot_id == "snap-v2"
    assert reports[0].snapshot_version == 2


def test_report_exposes_planned_actual_workload_and_capacity_utilization():
    report = build_therapist_daily_reports(
        (_snapshot("snap-v2", version=2, status=FinalizedSnapshotStatus.CURRENT_FINAL),),
        (_row("snap-v2"),),
    )[0]

    assert report.planned_occupied_timeslots == 5
    assert report.final_occupied_timeslots == 4
    assert report.planned_sessions == 5
    assert report.final_active_sessions == 4
    assert report.replacements_received == 1
    assert report.released_by_cancellation_or_no_show == 1
    assert report.planned_capacity_utilization == pytest.approx(5 / 6)
    assert report.actual_capacity_utilization == pytest.approx(4 / 6)
    assert report.session_delta == -1


def test_report_filters_therapist_and_date_range():
    next_day = date(2026, 10, 31)
    snapshots = (
        _snapshot("snap-a", version=1, status=FinalizedSnapshotStatus.CURRENT_FINAL),
        _snapshot(
            "snap-b",
            version=1,
            status=FinalizedSnapshotStatus.CURRENT_FINAL,
            service_date=next_day,
        ),
    )
    rows = (
        _row("snap-a", "T1"),
        _row("snap-a", "T2"),
        _row("snap-b", "T1", service_date=next_day),
    )

    reports = build_therapist_daily_reports(
        snapshots,
        rows,
        therapist_id="T1",
        start_date=next_day,
        end_date=next_day,
    )

    assert [(report.service_date, report.therapist_id) for report in reports] == [(next_day, "T1")]


def test_report_fails_closed_on_multiple_current_versions_for_same_day():
    snapshots = (
        _snapshot("snap-v1", version=1, status=FinalizedSnapshotStatus.CURRENT_FINAL),
        _snapshot("snap-v2", version=2, status=FinalizedSnapshotStatus.CURRENT_FINAL),
    )

    with pytest.raises(TherapistHistoryReportError, match="Multiple CURRENT_FINAL"):
        build_therapist_daily_reports(snapshots, (_row("snap-v2"),))


def test_report_fails_closed_on_duplicate_therapist_row_for_current_snapshot():
    snapshots = (_snapshot("snap-v2", version=2, status=FinalizedSnapshotStatus.CURRENT_FINAL),)

    with pytest.raises(TherapistHistoryReportError, match="Duplicate therapist/day history row"):
        build_therapist_daily_reports(snapshots, (_row("snap-v2"), _row("snap-v2")))


def test_report_fails_closed_if_current_snapshot_has_unresolved_sessions():
    snapshots = (
        _snapshot(
            "snap-v2",
            version=2,
            status=FinalizedSnapshotStatus.CURRENT_FINAL,
            unresolved_sessions=1,
        ),
    )

    with pytest.raises(TherapistHistoryReportError, match="unresolved sessions"):
        build_therapist_daily_reports(snapshots, (_row("snap-v2"),))
