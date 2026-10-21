from datetime import date, time

import pytest

from rehab_core.global_replacement_end_to_end import GlobalReplacementAssignmentResult
from rehab_core.models import ReplacementAssignment, Session
from rehab_excel.global_replacement_lifecycle_bridge import (
    GlobalReplacementLifecycleError,
    prepare_global_replacement_fact_batch,
)


TODAY = date(2026, 10, 21)


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=TODAY,
        start_time=time(hh, 0),
    )


def _replacement(
    replacement_id: str,
    target: Session,
    provider_id: str,
    hh: int,
) -> ReplacementAssignment:
    return ReplacementAssignment(
        replacement_id=replacement_id,
        target_session_id=target.session_id,
        patient_id=target.patient_id,
        original_therapist_id=target.therapist_id,
        replacement_therapist_id=provider_id,
        replacement_date=target.session_date,
        replacement_time=time(hh, 0),
        reason="global replacement plan",
    )


def test_prepares_global_assignments_for_existing_authoritative_fact_schema():
    a = _session("A", "P-A", "ABSENT-1", 9)
    b = _session("B", "P-B", "ABSENT-2", 10)
    result = GlobalReplacementAssignmentResult(
        assignments=(
            _replacement("R-A", a, "T1", 9),
            _replacement("R-B", b, "T2", 10),
        ),
        unresolved_request_ids=(),
    )

    batch = prepare_global_replacement_fact_batch(
        result,
        sessions=(a, b),
    )

    assert batch.unresolved_request_ids == ()
    assert [row.replacement_id for row in batch.fact_rows] == ["R-A", "R-B"]
    assert [row.target_session_id for row in batch.fact_rows] == ["A", "B"]
    assert all(row.status == "ACTIVE" for row in batch.fact_rows)


def test_partial_plan_preserves_unresolved_ids_without_inventing_facts():
    a = _session("A", "P-A", "ABSENT-1", 9)
    b = _session("B", "P-B", "ABSENT-2", 10)
    result = GlobalReplacementAssignmentResult(
        assignments=(_replacement("R-A", a, "T1", 9),),
        unresolved_request_ids=("B",),
    )

    batch = prepare_global_replacement_fact_batch(result, sessions=(a, b))

    assert [row.target_session_id for row in batch.fact_rows] == ["A"]
    assert batch.unresolved_request_ids == ("B",)


def test_rejects_conflict_with_existing_active_replacement_before_persistence():
    a = _session("A", "P-A", "ABSENT-1", 9)
    result = GlobalReplacementAssignmentResult(
        assignments=(_replacement("R-NEW", a, "T2", 10),),
        unresolved_request_ids=(),
    )
    existing = _replacement("R-OLD", a, "T1", 9)

    with pytest.raises(GlobalReplacementLifecycleError, match="already has an active replacement"):
        prepare_global_replacement_fact_batch(
            result,
            sessions=(a,),
            existing_active_replacements=(existing,),
        )


def test_allows_idempotent_existing_active_replacement():
    a = _session("A", "P-A", "ABSENT-1", 9)
    replacement = _replacement("R-A", a, "T1", 9)
    result = GlobalReplacementAssignmentResult(
        assignments=(replacement,),
        unresolved_request_ids=(),
    )

    batch = prepare_global_replacement_fact_batch(
        result,
        sessions=(a,),
        existing_active_replacements=(replacement,),
    )

    assert batch.fact_rows == ()
    assert batch.already_active_replacement_ids == ("R-A",)
