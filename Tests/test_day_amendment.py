from datetime import date, datetime, time, timezone

import pytest

from rehab_core.daily_state import build_daily_session_states
from rehab_core.day_amendment import DayAmendmentError, build_day_amendment_plan
from rehab_core.day_finalization import (
    FinalizedSnapshotStatus,
    build_finalized_day_snapshot,
)
from rehab_core.day_history_events import DayHistoryEventType
from rehab_core.models import AbsenceKind, DailyAbsence, Session


DAY = date(2026, 10, 30)
V1_AT = datetime(2026, 10, 30, 15, 0, tzinfo=timezone.utc)
STARTED_AT = datetime(2026, 10, 31, 9, 0, tzinfo=timezone.utc)
FINALIZED_AT = datetime(2026, 10, 31, 9, 15, tzinfo=timezone.utc)


def _session(session_id: str = "S1") -> Session:
    return Session(
        session_id=session_id,
        patient_id="P1",
        therapist_id="T1",
        session_date=DAY,
        start_time=time(8, 30),
        treatment="ΦΘ",
    )


def _prior_snapshot():
    return build_finalized_day_snapshot(
        build_daily_session_states([_session()]),
        snapshot_id="finalized:2026-10-30:v1",
        service_date=DAY,
        finalized_at=V1_AT,
        planned_session_count=1,
    )


def test_amendment_builds_version_chain_and_linked_events():
    prior = _prior_snapshot()
    states = build_daily_session_states(
        [_session()],
        absences=[
            DailyAbsence(
                absence_kind=AbsenceKind.PATIENT,
                subject_id="P1",
                absence_date=DAY,
                reason="Διόρθωση απουσίας",
            )
        ],
    )

    plan = build_day_amendment_plan(
        states,
        previous_snapshot=prior,
        replacement_snapshot_id="finalized:2026-10-30:v2",
        started_event_id="day-amendment-started:2026-10-30:v2",
        finalized_event_id="day-amendment-finalized:2026-10-30:v2",
        actor_or_source="operator",
        amendment_started_at=STARTED_AT,
        amendment_finalized_at=FINALIZED_AT,
        reason_code="PATIENT_ABSENCE_CORRECTION",
        summary="Correct patient attendance outcome",
    )

    assert plan.previous_snapshot == prior
    assert plan.superseded_snapshot.snapshot_id == prior.snapshot_id
    assert plan.superseded_snapshot.status == FinalizedSnapshotStatus.SUPERSEDED

    replacement = plan.replacement_snapshot
    assert replacement.status == FinalizedSnapshotStatus.CURRENT_FINAL
    assert replacement.version == 2
    assert replacement.supersedes_snapshot_id == prior.snapshot_id
    assert replacement.patient_absence_count == 1
    assert replacement.final_session_count == 0

    assert plan.started_event.event_type == DayHistoryEventType.DAY_AMENDMENT_STARTED
    assert plan.started_event.related_snapshot_id == prior.snapshot_id
    assert plan.started_event.reason_code == "PATIENT_ABSENCE_CORRECTION"

    assert plan.finalized_event.event_type == DayHistoryEventType.DAY_AMENDMENT_FINALIZED
    assert plan.finalized_event.related_event_id == plan.started_event.event_id
    assert plan.finalized_event.related_snapshot_id == replacement.snapshot_id
    assert prior.snapshot_id in plan.finalized_event.summary
    assert replacement.snapshot_id in plan.finalized_event.summary


def test_amendment_requires_current_final_previous_snapshot():
    prior = _prior_snapshot()
    superseded = prior.__class__(
        **{**prior.__dict__, "status": FinalizedSnapshotStatus.SUPERSEDED}
    )

    with pytest.raises(DayAmendmentError, match="CURRENT_FINAL"):
        build_day_amendment_plan(
            build_daily_session_states([_session()]),
            previous_snapshot=superseded,
            replacement_snapshot_id="finalized:2026-10-30:v2",
            started_event_id="start-v2",
            finalized_event_id="final-v2",
            actor_or_source="operator",
            amendment_started_at=STARTED_AT,
            amendment_finalized_at=FINALIZED_AT,
            reason_code="CORRECTION",
            summary="Correction",
        )


def test_amendment_rejects_reusing_snapshot_or_event_identity():
    prior = _prior_snapshot()
    states = build_daily_session_states([_session()])

    with pytest.raises(DayAmendmentError, match="new SnapshotId"):
        build_day_amendment_plan(
            states,
            previous_snapshot=prior,
            replacement_snapshot_id=prior.snapshot_id,
            started_event_id="start-v2",
            finalized_event_id="final-v2",
            actor_or_source="operator",
            amendment_started_at=STARTED_AT,
            amendment_finalized_at=FINALIZED_AT,
            reason_code="CORRECTION",
            summary="Correction",
        )

    with pytest.raises(DayAmendmentError, match="must be different"):
        build_day_amendment_plan(
            states,
            previous_snapshot=prior,
            replacement_snapshot_id="finalized:2026-10-30:v2",
            started_event_id="same-event",
            finalized_event_id="same-event",
            actor_or_source="operator",
            amendment_started_at=STARTED_AT,
            amendment_finalized_at=FINALIZED_AT,
            reason_code="CORRECTION",
            summary="Correction",
        )


def test_amendment_fails_closed_when_corrected_day_is_unresolved():
    prior = _prior_snapshot()
    states = build_daily_session_states(
        [_session("S-UNRESOLVED")],
        absences=[
            DailyAbsence(
                absence_kind=AbsenceKind.THERAPIST,
                subject_id="T1",
                absence_date=DAY,
                reason="Therapist absent",
            )
        ],
    )

    with pytest.raises(ValueError, match="unresolved sessions"):
        build_day_amendment_plan(
            states,
            previous_snapshot=prior,
            replacement_snapshot_id="finalized:2026-10-30:v2",
            started_event_id="start-v2",
            finalized_event_id="final-v2",
            actor_or_source="operator",
            amendment_started_at=STARTED_AT,
            amendment_finalized_at=FINALIZED_AT,
            reason_code="CORRECTION",
            summary="Correction",
        )


def test_amendment_validates_actor_reason_summary_and_time_order():
    prior = _prior_snapshot()
    states = build_daily_session_states([_session()])

    with pytest.raises(DayAmendmentError, match="actor_or_source"):
        build_day_amendment_plan(
            states,
            previous_snapshot=prior,
            replacement_snapshot_id="finalized:2026-10-30:v2",
            started_event_id="start-v2",
            finalized_event_id="final-v2",
            actor_or_source=" ",
            amendment_started_at=STARTED_AT,
            amendment_finalized_at=FINALIZED_AT,
            reason_code="CORRECTION",
            summary="Correction",
        )

    with pytest.raises(DayAmendmentError, match="reason_code"):
        build_day_amendment_plan(
            states,
            previous_snapshot=prior,
            replacement_snapshot_id="finalized:2026-10-30:v2",
            started_event_id="start-v2",
            finalized_event_id="final-v2",
            actor_or_source="operator",
            amendment_started_at=STARTED_AT,
            amendment_finalized_at=FINALIZED_AT,
            reason_code=" ",
            summary="Correction",
        )

    with pytest.raises(DayAmendmentError, match="summary"):
        build_day_amendment_plan(
            states,
            previous_snapshot=prior,
            replacement_snapshot_id="finalized:2026-10-30:v2",
            started_event_id="start-v2",
            finalized_event_id="final-v2",
            actor_or_source="operator",
            amendment_started_at=STARTED_AT,
            amendment_finalized_at=FINALIZED_AT,
            reason_code="CORRECTION",
            summary=" ",
        )

    with pytest.raises(DayAmendmentError, match="cannot be earlier"):
        build_day_amendment_plan(
            states,
            previous_snapshot=prior,
            replacement_snapshot_id="finalized:2026-10-30:v2",
            started_event_id="start-v2",
            finalized_event_id="final-v2",
            actor_or_source="operator",
            amendment_started_at=FINALIZED_AT,
            amendment_finalized_at=STARTED_AT,
            reason_code="CORRECTION",
            summary="Correction",
        )
