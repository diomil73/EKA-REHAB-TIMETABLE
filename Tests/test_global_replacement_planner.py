from datetime import time
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_core.global_replacement_planner import (
    GlobalReplacementCandidate,
    GlobalReplacementRequest,
    plan_global_replacements,
)


def _candidate(
    request_id: str,
    provider_id: str,
    hh: int,
    *,
    original: bool = False,
    load: int = 0,
    adjacency_penalty: int = 0,
) -> GlobalReplacementCandidate:
    return GlobalReplacementCandidate(
        request_id=request_id,
        provider_id=provider_id,
        replacement_time=time(hh, 0),
        provider_active_timeslots=load,
        is_original_time=original,
        adjacency_penalty=adjacency_penalty,
    )


def test_protects_unique_slot_to_maximize_complete_assignment():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P-A",
            candidates=(
                _candidate("A", "T1", 9, original=True),
                _candidate("A", "T1", 10),
                _candidate("A", "T1", 11),
            ),
        ),
        GlobalReplacementRequest(
            request_id="B",
            patient_id="P-B",
            candidates=(_candidate("B", "T1", 9, original=True),),
        ),
    )

    plan = plan_global_replacements(requests)

    assert plan.unresolved_request_ids == ()
    by_request = {item.request_id: item for item in plan.assignments}
    assert by_request["B"].replacement_time == time(9, 0)
    assert by_request["A"].replacement_time == time(10, 0)


def test_maximizes_number_placed_before_preserving_original_time():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P-A",
            candidates=(
                _candidate("A", "T1", 9, original=True),
                _candidate("A", "T2", 10),
            ),
        ),
        GlobalReplacementRequest(
            request_id="B",
            patient_id="P-B",
            candidates=(_candidate("B", "T1", 9, original=True),),
        ),
    )

    plan = plan_global_replacements(requests)

    assert len(plan.assignments) == 2
    assert plan.unresolved_request_ids == ()


def test_same_patient_cannot_receive_two_replacements_at_same_time():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P1",
            candidates=(_candidate("A", "T1", 10),),
        ),
        GlobalReplacementRequest(
            request_id="B",
            patient_id="P1",
            candidates=(
                _candidate("B", "T2", 10),
                _candidate("B", "T2", 11),
            ),
        ),
    )

    plan = plan_global_replacements(requests)

    assert len(plan.assignments) == 2
    assert {item.replacement_time for item in plan.assignments} == {
        time(10, 0),
        time(11, 0),
    }


def test_prefers_reclining_adjacency_over_original_time_when_both_place_all():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P1",
            candidates=(
                _candidate("A", "T1", 9, original=True, adjacency_penalty=60),
                _candidate("A", "T1", 10, adjacency_penalty=0),
            ),
        ),
    )

    plan = plan_global_replacements(requests)

    assert plan.assignments[0].replacement_time == time(10, 0)


def test_prefers_better_final_workload_balance_when_higher_priorities_tie():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P1",
            candidates=(
                _candidate("A", "BUSY", 10, load=5),
                _candidate("A", "LIGHT", 10, load=3),
            ),
        ),
    )

    plan = plan_global_replacements(requests)

    assert plan.assignments[0].provider_id == "LIGHT"


def test_avoids_1300_only_as_weak_tie_breaker():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P1",
            candidates=(
                _candidate("A", "T1", 13),
                _candidate("A", "T1", 12),
            ),
        ),
    )

    plan = plan_global_replacements(requests)

    assert plan.assignments[0].replacement_time == time(12, 0)


def test_uses_1300_instead_of_leaving_patient_unresolved():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P1",
            candidates=(_candidate("A", "T1", 13),),
        ),
    )

    plan = plan_global_replacements(requests)

    assert len(plan.assignments) == 1
    assert plan.assignments[0].replacement_time == time(13, 0)
    assert plan.unresolved_request_ids == ()


def test_provider_remaining_capacity_is_respected_across_global_plan():
    requests = (
        GlobalReplacementRequest(
            request_id="A",
            patient_id="P1",
            candidates=(_candidate("A", "T1", 10),),
        ),
        GlobalReplacementRequest(
            request_id="B",
            patient_id="P2",
            candidates=(
                _candidate("B", "T1", 11),
                _candidate("B", "T2", 11),
            ),
        ),
    )

    plan = plan_global_replacements(
        requests,
        provider_capacity_remaining={"T1": 1, "T2": 1},
    )

    assert len(plan.assignments) == 2
    assert {item.provider_id for item in plan.assignments} == {"T1", "T2"}
