"""Read-only synthetic reconciliation of finalized day snapshots and history."""
from dataclasses import replace
from datetime import date, datetime, time, timedelta, timezone
from hashlib import sha256

import pytest
from openpyxl import Workbook, load_workbook

from rehab_core.daily_state import DailySessionStatus as Status
from rehab_core.day_finalization import FinalizedDaySnapshot, FinalizedSnapshotStatus
from rehab_core.day_history_events import DayHistoryEvent, DayHistoryEventType
from rehab_core.models import SessionCancellationKind
from rehab_core.patient_session_history import PatientSessionHistoryRow
from rehab_core.therapist_day_history import TherapistDayHistoryRow
from rehab_excel.finalized_history_consistency import (
    FinalizedHistoryConsistencyError,
    verify_current_finalized_day,
    verify_finalized_history_bundle,
)
from rehab_excel.finalized_day_snapshots import (
    HEADERS as SNAPSHOT_HEADERS,
    _values_for_com as snapshot_values,
)
from rehab_excel.patient_session_history_store import (
    HEADERS as PATIENT_HEADERS,
    values_for_payload,
)
from rehab_excel.therapist_day_history_store import (
    HEADERS as THERAPIST_HEADERS,
    _values_for_com as therapist_values,
)
from rehab_excel.day_history_event_store import (
    HEADERS as EVENT_HEADERS,
    _values_for_com as event_values,
)


DAY = date(2026, 11, 24)
STAMP = datetime(2026, 11, 24, 18, 0, tzinfo=timezone(timedelta(hours=2)))


def _bundle(*, version=1):
    sid = f"fictional-final:{DAY}:v{version}"
    snapshot = FinalizedDaySnapshot(
        snapshot_id=sid, service_date=DAY, version=version,
        status=FinalizedSnapshotStatus.CURRENT_FINAL, finalized_at=STAMP,
        supersedes_snapshot_id=None if version == 1 else "fictional-final:prior",
        unresolved_sessions=0, planned_session_count=5, final_session_count=2,
        patient_absence_count=1, cancellation_count=2, no_show_count=1,
        replacement_count=1, treatment_suspension_count=1,
        approved_exception_count=0,
    )

    def patient(index, status, *, reason=None, kind=None, replacement=None):
        return PatientSessionHistoryRow(
            snapshot_id=sid, service_date=DAY, snapshot_version=version,
            session_id=f"FICTIONAL-S{index}", patient_id=f"FICTIONAL-P{index}",
            status=status, original_therapist_id="FICTIONAL-T1",
            original_time=time(8 + index, 0),
            effective_therapist_id=(
                "FICTIONAL-T2" if status == Status.REPLACED else
                "FICTIONAL-T1" if status == Status.ACTIVE else None
            ),
            effective_time=time(8 + index, 0) if status in (
                Status.ACTIVE, Status.REPLACED,
            ) else None,
            replacement_id=replacement, cancellation_id=(
                f"FICTIONAL-C{index}" if kind else None
            ),
            cancellation_kind=kind, reason=reason,
        )

    patients = (
        patient(0, Status.ACTIVE),
        patient(1, Status.REPLACED, replacement="FICTIONAL-R1"),
        patient(2, Status.PATIENT_ABSENT, reason="FICTIONAL_ABSENCE"),
        patient(3, Status.CANCELLED, kind=SessionCancellationKind.PATIENT_NO_SHOW.value),
        patient(4, Status.CANCELLED, kind=SessionCancellationKind.TREATMENT_SUSPENDED.value),
    )
    therapists = (
        TherapistDayHistoryRow(
            snapshot_id=sid, service_date=DAY, therapist_id="FICTIONAL-T1",
            capacity_timeslots=6, planned_occupied_timeslots=5,
            planned_sessions=5, final_occupied_timeslots=1,
            final_active_sessions=1, replacements_received=0,
            released_by_cancellation_or_no_show=2,
            treatment_suspensions_affecting_load=1,
        ),
        TherapistDayHistoryRow(
            snapshot_id=sid, service_date=DAY, therapist_id="FICTIONAL-T2",
            capacity_timeslots=6, planned_occupied_timeslots=0,
            planned_sessions=0, final_occupied_timeslots=1,
            final_active_sessions=1, replacements_received=1,
            released_by_cancellation_or_no_show=0,
            treatment_suspensions_affecting_load=0,
        ),
    )
    event = DayHistoryEvent(
        event_id=f"fictional-event:v{version}", service_date=DAY,
        event_type=(
            DayHistoryEventType.DAY_FINALIZED if version == 1 else
            DayHistoryEventType.DAY_AMENDMENT_FINALIZED
        ),
        recorded_at=STAMP, actor_or_source="FICTIONAL_OPERATOR",
        related_snapshot_id=sid, summary="Synthetic day finalization",
    )
    return snapshot, patients, therapists, (event,)


def _book(path, snapshot, patients, therapists, events):
    wb = Workbook()
    ws = wb.active
    ws.title = "FINALIZED_DAY_SNAPSHOTS"
    ws.append(SNAPSHOT_HEADERS)
    ws.append(snapshot_values(snapshot))
    p = wb.create_sheet("PATIENT_SESSION_HISTORY")
    p.append(PATIENT_HEADERS)
    for row in patients:
        p.append(values_for_payload(row))
    t = wb.create_sheet("THERAPIST_DAY_HISTORY")
    t.append(THERAPIST_HEADERS)
    for row in therapists:
        t.append(therapist_values(row))
    e = wb.create_sheet("DAY_HISTORY_EVENTS")
    e.append(EVENT_HEADERS)
    for row in events:
        e.append(event_values(row))
    wb.save(path)
    wb.close()
    return path


def test_consistent_synthetic_bundle_has_matching_counts_without_attendance_guess():
    report = verify_finalized_history_bundle(*_bundle())
    assert report.patient_session_rows == 5
    assert report.therapist_day_rows == 2
    assert report.finalization_event_count == 1
    assert report.version == 1
    assert report.snapshot_sha256 is None


def test_amended_version_requires_amendment_finalized_event():
    report = verify_finalized_history_bundle(*_bundle(version=2))
    assert report.version == 2


@pytest.mark.parametrize("field", [
    "planned_session_count", "final_session_count", "patient_absence_count",
    "cancellation_count", "no_show_count", "replacement_count",
    "treatment_suspension_count",
])
def test_all_snapshot_aggregates_are_independently_recomputed(field):
    snapshot, patients, therapists, events = _bundle()
    incorrect = replace(snapshot, **{field: getattr(snapshot, field) + 1})
    with pytest.raises(FinalizedHistoryConsistencyError, match="count|totals"):
        verify_finalized_history_bundle(incorrect, patients, therapists, events)


def test_patient_wrong_snapshot_date_or_version_is_rejected():
    snapshot, patients, therapists, events = _bundle()
    for change in (
        {"snapshot_id": "wrong-snapshot"},
        {"service_date": date(2026, 11, 25)},
        {"snapshot_version": 2},
    ):
        modified = (replace(patients[0], **change), *patients[1:])
        with pytest.raises(FinalizedHistoryConsistencyError, match="snapshot/date/version"):
            verify_finalized_history_bundle(snapshot, modified, therapists, events)


def test_duplicate_session_identity_is_rejected_case_insensitively():
    snapshot, patients, therapists, events = _bundle()
    modified = (*patients[:-1], replace(patients[-1], session_id=patients[0].session_id.lower()))
    with pytest.raises(FinalizedHistoryConsistencyError, match="Duplicate"):
        verify_finalized_history_bundle(snapshot, modified, therapists, events)


def test_therapist_mismatch_duplicate_and_counter_drift_are_rejected():
    snapshot, patients, therapists, events = _bundle()
    for replacement in (
        replace(therapists[0], snapshot_id="other"),
        replace(therapists[0], service_date=date(2026, 11, 25)),
        replace(therapists[0], final_occupied_timeslots=8),
    ):
        with pytest.raises(FinalizedHistoryConsistencyError):
            verify_finalized_history_bundle(snapshot, patients, (replacement, therapists[1]), events)
    dup = replace(therapists[1], therapist_id=therapists[0].therapist_id.lower())
    with pytest.raises(FinalizedHistoryConsistencyError, match="Duplicate"):
        verify_finalized_history_bundle(snapshot, patients, (therapists[0], dup), events)


def test_finalization_event_is_mandatory_unique_and_version_matched():
    snapshot, patients, therapists, events = _bundle()
    for altered in (
        (),
        (events[0], replace(events[0], event_id="FICTIONAL-DUPLICATE")),
        (replace(events[0], service_date=date(2026, 11, 25)),),
        (replace(events[0], event_type=DayHistoryEventType.DAY_AMENDMENT_FINALIZED),),
    ):
        with pytest.raises(FinalizedHistoryConsistencyError):
            verify_finalized_history_bundle(snapshot, patients, therapists, altered)


def test_wrong_status_and_unresolved_therapist_absence_fail_closed():
    snapshot, patients, therapists, events = _bundle()
    with pytest.raises(FinalizedHistoryConsistencyError, match="CURRENT_FINAL"):
        verify_finalized_history_bundle(
            replace(snapshot, status=FinalizedSnapshotStatus.SUPERSEDED),
            patients, therapists, events,
        )
    with pytest.raises(FinalizedHistoryConsistencyError, match="Unresolved"):
        verify_finalized_history_bundle(
            snapshot,
            (replace(patients[0], status=Status.THERAPIST_ABSENT), *patients[1:]),
            therapists, events,
        )


def test_real_xlsx_readonly_reconciliation_does_not_mutate_workbook(tmp_path):
    snapshot, patients, therapists, events = _bundle()
    path = _book(tmp_path / "fictional-history.xlsx", snapshot, patients, therapists, events)
    before = sha256(path.read_bytes()).hexdigest()
    report = verify_current_finalized_day(path, service_date=DAY)
    assert report.snapshot_id == snapshot.snapshot_id
    assert report.patient_session_rows == 5
    assert report.therapist_day_rows == 2
    assert report.snapshot_sha256 == before
    assert sha256(path.read_bytes()).hexdigest() == before


def test_missing_history_sheet_and_stale_summary_refuse_without_changes(tmp_path):
    snapshot, patients, therapists, events = _bundle()
    path = _book(tmp_path / "fictional-history.xlsx", snapshot, patients, therapists, events)
    wb = load_workbook(path)
    wb["FINALIZED_DAY_SNAPSHOTS"]["I2"] = 999
    wb.save(path)
    wb.close()
    unchanged = path.read_bytes()
    with pytest.raises(FinalizedHistoryConsistencyError, match="totals"):
        verify_current_finalized_day(path, service_date=DAY)
    assert path.read_bytes() == unchanged

    wb = load_workbook(path)
    wb["FINALIZED_DAY_SNAPSHOTS"]["I2"] = snapshot.final_session_count
    del wb["DAY_HISTORY_EVENTS"]
    wb.save(path)
    wb.close()
    unchanged = path.read_bytes()
    with pytest.raises(FinalizedHistoryConsistencyError, match="finalization event"):
        verify_current_finalized_day(path, service_date=DAY)
    assert path.read_bytes() == unchanged


def test_unfinalized_or_nonexistent_date_never_counts_as_empty_success(tmp_path):
    snapshot, patients, therapists, events = _bundle()
    path = _book(tmp_path / "fictional-history.xlsx", snapshot, patients, therapists, events)
    with pytest.raises(FinalizedHistoryConsistencyError, match="CURRENT_FINAL"):
        verify_current_finalized_day(path, service_date=date(2026, 11, 25))
