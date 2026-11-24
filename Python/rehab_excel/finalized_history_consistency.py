"""Read-only cross-sheet consistency audit for a CURRENT_FINAL day.

This is a necessary history proof, not a sufficient clinical authorization.
It independently recomputes session-count fields from durable patient history,
verifies snapshot/date/version identity in therapist rows, and links the
finalization event to the exact snapshot. It never infers real therapist
attendance from scheduled sessions, nor writes back to a workbook.

Use this before treating finalized patient/therapist reports as verified.
Full workbook transaction Save/Close/Reopen and operator approval are separate.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from rehab_core.daily_state import DailySessionStatus
from rehab_core.day_finalization import FinalizedDaySnapshot, FinalizedSnapshotStatus
from rehab_core.day_history_events import DayHistoryEvent, DayHistoryEventType
from rehab_core.models import SessionCancellationKind
from rehab_core.patient_session_history import PatientSessionHistoryRow
from rehab_core.therapist_day_history import TherapistDayHistoryRow

from .authoritative_commit import file_sha256
from .day_history_event_store import read_day_history_events
from .finalized_day_snapshots import read_finalized_day_snapshots
from .patient_session_history_store import read_patient_session_history_rows
from .therapist_day_history_store import read_therapist_day_history_rows


class FinalizedHistoryConsistencyError(RuntimeError):
    """Independent historical facts do not agree; never approve this day."""


@dataclass(frozen=True)
class FinalizedHistoryConsistencyReport:
    snapshot_id: str
    service_date: date
    version: int
    patient_session_rows: int
    therapist_day_rows: int
    finalization_event_count: int
    snapshot_sha256: str | None = None


def verify_finalized_history_bundle(
    snapshot: FinalizedDaySnapshot,
    patient_rows: tuple[PatientSessionHistoryRow, ...],
    therapist_rows: tuple[TherapistDayHistoryRow, ...],
    events: tuple[DayHistoryEvent, ...],
) -> FinalizedHistoryConsistencyReport:
    """Compare independently persisted finalized history records, fail closed.

    Patient rows represent one row per immutable base session. They are NOT
    therapist attendance records; therapist rows reflect workload only.
    """
    if snapshot.status != FinalizedSnapshotStatus.CURRENT_FINAL:
        raise FinalizedHistoryConsistencyError("Historical snapshot is not CURRENT_FINAL")
    if snapshot.version < 1 or not snapshot.snapshot_id or snapshot.unresolved_sessions != 0:
        raise FinalizedHistoryConsistencyError("Historical snapshot identity or state invalid")
    if snapshot.planned_session_count is None:
        raise FinalizedHistoryConsistencyError("Planned session count unavailable")
    if snapshot.planned_session_count != len(patient_rows):
        raise FinalizedHistoryConsistencyError("Planned session count differs from patient history")
    if not therapist_rows:
        raise FinalizedHistoryConsistencyError("Finalized therapist/day history missing")

    seen_sessions: set[str] = set()
    for row in patient_rows:
        if (row.snapshot_id != snapshot.snapshot_id
                or row.service_date != snapshot.service_date
                or row.snapshot_version != snapshot.version):
            raise FinalizedHistoryConsistencyError("Patient history snapshot/date/version mismatch")
        key = row.session_id.strip().casefold()
        if not key or key in seen_sessions or not row.patient_id.strip():
            raise FinalizedHistoryConsistencyError("Duplicate or invalid patient history session")
        seen_sessions.add(key)
        if row.status == DailySessionStatus.THERAPIST_ABSENT:
            raise FinalizedHistoryConsistencyError("Unresolved therapist absence in finalized history")

    kinds = SessionCancellationKind
    def count_status(status: DailySessionStatus) -> int:
        return sum(row.status == status for row in patient_rows)

    final_sessions = count_status(DailySessionStatus.ACTIVE) + count_status(DailySessionStatus.REPLACED)
    canceled = tuple(row for row in patient_rows if row.status == DailySessionStatus.CANCELLED)
    no_shows = sum(row.cancellation_kind == kinds.PATIENT_NO_SHOW.value for row in canceled)
    suspensions = sum(row.cancellation_kind == kinds.TREATMENT_SUSPENDED.value for row in canceled)

    if (
        snapshot.final_session_count != final_sessions
        or snapshot.patient_absence_count != count_status(DailySessionStatus.PATIENT_ABSENT)
        or snapshot.cancellation_count != len(canceled)
        or snapshot.no_show_count != no_shows
        or snapshot.replacement_count != count_status(DailySessionStatus.REPLACED)
        or snapshot.treatment_suspension_count != suspensions
    ):
        raise FinalizedHistoryConsistencyError("Snapshot totals differ from patient history")

    if snapshot.approved_exception_count < 0:
        raise FinalizedHistoryConsistencyError("Snapshot exception count invalid")

    seen_therapists: set[str] = set()
    for row in therapist_rows:
        if row.snapshot_id != snapshot.snapshot_id or row.service_date != snapshot.service_date:
            raise FinalizedHistoryConsistencyError("Therapist history snapshot/date mismatch")
        key = row.therapist_id.strip().casefold()
        if not key or key in seen_therapists:
            raise FinalizedHistoryConsistencyError("Duplicate or invalid therapist history")
        seen_therapists.add(key)
        if (
            row.planned_occupied_timeslots > row.planned_sessions
            or row.final_occupied_timeslots > row.final_active_sessions
            or row.replacements_received > row.final_active_sessions
        ):
            raise FinalizedHistoryConsistencyError("Therapist workload counters inconsistent")
    # Do NOT assert sum of therapist final_active_sessions == performed patient
    # sessions: student-delivered replacements are intentionally excluded.

    matching_events = tuple(
        event for event in events
        if event.related_snapshot_id == snapshot.snapshot_id
        and event.event_type in {
            DayHistoryEventType.DAY_FINALIZED,
            DayHistoryEventType.DAY_AMENDMENT_FINALIZED,
        }
    )
    if len(matching_events) != 1:
        raise FinalizedHistoryConsistencyError("Exactly one finalization event required")
    event = matching_events[0]
    if event.service_date != snapshot.service_date:
        raise FinalizedHistoryConsistencyError("Finalization event date mismatch")
    if (
        (snapshot.version == 1 and event.event_type != DayHistoryEventType.DAY_FINALIZED)
        or (snapshot.version > 1 and event.event_type != DayHistoryEventType.DAY_AMENDMENT_FINALIZED)
    ):
        raise FinalizedHistoryConsistencyError("Finalization event type/version mismatch")
    if any(
        row.related_snapshot_id == snapshot.snapshot_id and row != event
        for row in events
    ):
        # A snapshot may have separate amendment-start events. Their history is
        # allowed, but they do not substitute for its required final event.
        pass

    return FinalizedHistoryConsistencyReport(
        snapshot_id=snapshot.snapshot_id,
        service_date=snapshot.service_date,
        version=snapshot.version,
        patient_session_rows=len(patient_rows),
        therapist_day_rows=len(therapist_rows),
        finalization_event_count=1,
    )


def verify_current_finalized_day(
    workbook_path: str | Path,
    *,
    service_date: date,
) -> FinalizedHistoryConsistencyReport:
    """Reopen the saved workbook read-only across all four history sheets.

    A complete-file SHA checked before and after repeated independent reads
    prevents mixing two consecutive workbook revisions during this audit.
    No missing sheet or missing day is interpreted as an empty finalized day.
    """
    path = Path(workbook_path).resolve()
    if not path.is_file() or path.suffix.casefold() not in {".xlsm", ".xlsx"}:
        raise FinalizedHistoryConsistencyError("Existing Excel workbook required")
    sha_before = file_sha256(path)
    try:
        snapshots = read_finalized_day_snapshots(
            path, service_date=service_date, current_only=True,
        )
        if len(snapshots) != 1:
            raise FinalizedHistoryConsistencyError("Exactly one CURRENT_FINAL snapshot required")
        snapshot = snapshots[0]
        patients = read_patient_session_history_rows(
            path, snapshot_id=snapshot.snapshot_id,
        )
        therapists = read_therapist_day_history_rows(
            path, snapshot_id=snapshot.snapshot_id,
        )
        events = read_day_history_events(path, service_date=service_date)
        report = verify_finalized_history_bundle(
            snapshot, patients, therapists, events,
        )
    except FinalizedHistoryConsistencyError:
        raise
    except Exception as exc:
        raise FinalizedHistoryConsistencyError("Historical workbook readback invalid") from exc
    if file_sha256(path) != sha_before:
        raise FinalizedHistoryConsistencyError("Workbook changed during historical audit")
    return FinalizedHistoryConsistencyReport(
        snapshot_id=report.snapshot_id,
        service_date=report.service_date,
        version=report.version,
        patient_session_rows=report.patient_session_rows,
        therapist_day_rows=report.therapist_day_rows,
        finalization_event_count=report.finalization_event_count,
        snapshot_sha256=sha_before,
    )
