from dataclasses import FrozenInstanceError
from datetime import date, datetime, time, timezone

import pytest

from rehab_core.day_finalization import (
    DayFinalizationError,
    FinalizedSnapshotStatus,
    build_finalized_day_snapshot,
)
from rehab_core.daily_state import build_daily_session_states
from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    DailySessionCancellation,
    ReplacementAssignment,
    Session,
    SessionCancellationKind,
)
from rehab_core.orphan_resolution import create_treatment_suspension


DAY = date(2026, 10, 26)
FINALIZED_AT = datetime(2026, 10, 26, 15, 30, tzinfo=timezone.utc)


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int, mm: int) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=time(hh, mm),
        treatment="ΦΘ",
    )


def test_finalization_fails_closed_when_any_session_is_unresolved():
    unresolved = _session("S-ORPH", "P1", "T1", 8, 30)
    states = build_daily_session_states(
        [unresolved],
        absences=[
            DailyAbsence(
                absence_kind=AbsenceKind.THERAPIST,
                subject_id="T1",
                absence_date=DAY,
                reason="Απουσία θεραπευτή",
            )
        ],
    )

    with pytest.raises(DayFinalizationError, match="S-ORPH"):
        build_finalized_day_snapshot(
            states,
            snapshot_id="snap-2026-10-26-v1",
            service_date=DAY,
            finalized_at=FINALIZED_AT,
        )


def test_finalized_snapshot_freezes_effective_daily_outcomes():
    active = _session("S-ACT", "P-ACT", "T1", 8, 30)
    replaced = _session("S-REP", "P-REP", "T2", 9, 15)
    absent = _session("S-ABS", "P-ABS", "T3", 10, 0)
    no_show = _session("S-NS", "P-NS", "T4", 10, 45)
    suspended = _session("S-SUS", "P-SUS", "T5", 11, 30)

    states = build_daily_session_states(
        [active, replaced, absent, no_show, suspended],
        replacements=[
            ReplacementAssignment(
                replacement_id="R1",
                target_session_id=replaced.session_id,
                patient_id=replaced.patient_id,
                original_therapist_id=replaced.therapist_id,
                replacement_therapist_id="T9",
                replacement_date=DAY,
                replacement_time=time(12, 15),
            )
        ],
        absences=[
            DailyAbsence(
                absence_kind=AbsenceKind.PATIENT,
                subject_id=absent.patient_id,
                absence_date=DAY,
                reason="Απουσία ασθενή",
            )
        ],
        cancellations=[
            DailySessionCancellation(
                cancellation_id="C-NS",
                target_session_id=no_show.session_id,
                cancellation_date=DAY,
                kind=SessionCancellationKind.PATIENT_NO_SHOW,
                reason="Δεν προσήλθε",
            ),
            create_treatment_suspension(suspended),
        ],
    )

    snapshot = build_finalized_day_snapshot(
        states,
        snapshot_id="snap-2026-10-26-v1",
        service_date=DAY,
        finalized_at=FINALIZED_AT,
        planned_session_count=5,
        approved_exception_count=1,
    )

    assert snapshot.status == FinalizedSnapshotStatus.CURRENT_FINAL
    assert snapshot.version == 1
    assert snapshot.unresolved_sessions == 0
    assert snapshot.planned_session_count == 5
    assert snapshot.final_session_count == 2
    assert snapshot.patient_absence_count == 1
    assert snapshot.cancellation_count == 2
    assert snapshot.no_show_count == 1
    assert snapshot.replacement_count == 1
    assert snapshot.treatment_suspension_count == 1
    assert snapshot.approved_exception_count == 1

    with pytest.raises(FrozenInstanceError):
        snapshot.final_session_count = 999  # type: ignore[misc]


def test_finalization_rejects_states_from_another_service_date():
    wrong_day = Session(
        session_id="S-WRONG-DATE",
        patient_id="P1",
        therapist_id="T1",
        session_date=date(2026, 10, 27),
        start_time=time(8, 30),
        treatment="ΦΘ",
    )
    states = build_daily_session_states([wrong_day])

    with pytest.raises(DayFinalizationError, match="S-WRONG-DATE"):
        build_finalized_day_snapshot(
            states,
            snapshot_id="snap-2026-10-26-v1",
            service_date=DAY,
            finalized_at=FINALIZED_AT,
        )


def test_finalization_validates_header_inputs_before_persistence():
    with pytest.raises(DayFinalizationError, match="snapshot_id"):
        build_finalized_day_snapshot(
            [],
            snapshot_id="   ",
            service_date=DAY,
            finalized_at=FINALIZED_AT,
        )

    with pytest.raises(DayFinalizationError, match="version"):
        build_finalized_day_snapshot(
            [],
            snapshot_id="snap",
            service_date=DAY,
            finalized_at=FINALIZED_AT,
            version=0,
        )

    with pytest.raises(DayFinalizationError, match="approved_exception_count"):
        build_finalized_day_snapshot(
            [],
            snapshot_id="snap",
            service_date=DAY,
            finalized_at=FINALIZED_AT,
            approved_exception_count=-1,
        )
