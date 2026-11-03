from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from rehab_core.daily_state import DailySessionStatus
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots
from rehab_excel.patient_session_history_store import read_current_patient_session_history_rows
from rehab_excel.therapist_day_history_store import read_therapist_day_history_rows


class HistoricalStatisticsError(ValueError):
    """Fail-closed if finalized history is incomplete or inconsistent."""


@dataclass(frozen=True)
class DailyStatistics:
    service_date: date
    snapshot_id: str
    version: int
    scheduled: int
    performed: int
    patient_absences: int
    cancellations: int
    replacements: int
    therapists: int


@dataclass(frozen=True)
class TherapistStatistics:
    therapist_id: str
    days: int
    planned_sessions: int
    performed_sessions: int
    replacements_received: int
    released_by_cancellation_or_no_show: int
    treatment_suspensions: int


def summarize_final_history(snapshots, patient_rows, therapist_rows, *, start: date | None = None, end: date | None = None):
    """Build consistent daily and therapist summaries from CURRENT_FINAL data.

    The snapshots argument must contain only current snapshots. Caller responsibility
    is explicit; the workbook reader below selects CURRENT_FINAL authoritatively.
    """
    if start is not None and end is not None and start > end:
        raise HistoricalStatisticsError("start date cannot be after end date")

    selected = tuple(
        s for s in snapshots
        if (start is None or s.service_date >= start)
        and (end is None or s.service_date <= end)
    )
    ids = set()
    dates = set()
    for snap in selected:
        if snap.snapshot_id in ids or snap.service_date in dates:
            raise HistoricalStatisticsError("Duplicate current snapshot/date")
        ids.add(snap.snapshot_id)
        dates.add(snap.service_date)
        if snap.status.value != "CURRENT_FINAL" or snap.unresolved_sessions != 0:
            raise HistoricalStatisticsError("Non-current or unresolved snapshot")

    patients = defaultdict(list)
    for row in patient_rows:
        if row.snapshot_id not in ids:
            continue
        patients[row.snapshot_id].append(row)

    therapists = defaultdict(list)
    for row in therapist_rows:
        if row.snapshot_id not in ids:
            continue
        therapists[row.snapshot_id].append(row)

    daily = []
    totals = defaultdict(lambda: [0, 0, 0, 0, 0, 0])
    for snap in sorted(selected, key=lambda x: x.service_date):
        rows = patients[snap.snapshot_id]
        trows = therapists[snap.snapshot_id]
        seen_sessions = set()
        for row in rows:
            if row.service_date != snap.service_date or row.snapshot_version != snap.version:
                raise HistoricalStatisticsError("Patient history date/version mismatch")
            key = row.session_id.casefold()
            if key in seen_sessions:
                raise HistoricalStatisticsError("Duplicate patient session in current snapshot")
            seen_sessions.add(key)
            if row.status == DailySessionStatus.THERAPIST_ABSENT:
                raise HistoricalStatisticsError("Unresolved patient session")
        performed = sum(r.status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED) for r in rows)
        absent = sum(r.status == DailySessionStatus.PATIENT_ABSENT for r in rows)
        cancelled = sum(r.status == DailySessionStatus.CANCELLED for r in rows)
        replaced = sum(r.status == DailySessionStatus.REPLACED for r in rows)
        if snap.planned_session_count is not None and len(rows) != snap.planned_session_count:
            raise HistoricalStatisticsError(f"Missing patient history rows: {snap.snapshot_id}")
        expected = (snap.final_session_count, snap.patient_absence_count, snap.cancellation_count, snap.replacement_count)
        actual = (performed, absent, cancelled, replaced)
        if actual != expected:
            raise HistoricalStatisticsError(f"Snapshot totals disagree with patient history: {snap.snapshot_id}")

        therapist_ids = set()
        for row in trows:
            if row.service_date != snap.service_date:
                raise HistoricalStatisticsError("Therapist history date mismatch")
            if row.therapist_id.casefold() in therapist_ids:
                raise HistoricalStatisticsError("Duplicate therapist in current snapshot")
            therapist_ids.add(row.therapist_id.casefold())
            counters = totals[row.therapist_id]
            counters[0] += 1
            counters[1] += row.planned_sessions
            counters[2] += row.final_active_sessions
            counters[3] += row.replacements_received
            counters[4] += row.released_by_cancellation_or_no_show
            counters[5] += row.treatment_suspensions_affecting_load
        if not trows and rows:
            raise HistoricalStatisticsError("Missing therapist history rows")

        daily.append(DailyStatistics(
            service_date=snap.service_date,
            snapshot_id=snap.snapshot_id,
            version=snap.version,
            scheduled=len(rows),
            performed=performed,
            patient_absences=absent,
            cancellations=cancelled,
            replacements=replaced,
            therapists=len(trows),
        ))
    therapist_summaries = tuple(
        TherapistStatistics(
            therapist_id=key, days=value[0], planned_sessions=value[1],
            performed_sessions=value[2], replacements_received=value[3],
            released_by_cancellation_or_no_show=value[4],
            treatment_suspensions=value[5],
        )
        for key, value in sorted(totals.items(), key=lambda item: item[0].casefold())
    )
    return tuple(daily), therapist_summaries


def read_historical_statistics(workbook_path: str | Path, *, start: date | None = None, end: date | None = None):
    """Read-only, version-scoped backend. No Excel COM or workbook mutations."""
    snapshots = read_finalized_day_snapshots(workbook_path, current_only=True)
    patients = read_current_patient_session_history_rows(workbook_path)
    current_ids = {snap.snapshot_id for snap in snapshots}
    therapists = tuple(row for row in read_therapist_day_history_rows(workbook_path) if row.snapshot_id in current_ids)
    return summarize_final_history(snapshots, patients, therapists, start=start, end=end)
