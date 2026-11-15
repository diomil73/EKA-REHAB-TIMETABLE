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


def _candidate(request_id: str, provider_id: str, hh: int, mm: int = 0):
    return GlobalReplacementCandidate(
        request_id=request_id,
        provider_id=provider_id,
        replacement_time=time(hh, mm),
        provider_active_timeslots=3,
        is_original_time=(hh, mm) == (10, 0),
    )


def test_six_request_multi_absence_search_remains_practical_and_complete():
    slots = ((8, 30), (9, 15), (10, 0), (10, 45), (11, 30), (12, 15), (13, 0))
    providers = ("T1", "T2", "T3", "T4", "T5", "T6")

    requests = tuple(
        GlobalReplacementRequest(
            request_id=f"R{request_index}",
            patient_id=f"P{request_index}",
            candidates=tuple(
                _candidate(f"R{request_index}", provider_id, hh, mm)
                for provider_id in providers
                for hh, mm in slots
            ),
        )
        for request_index in range(1, 7)
    )

    plan = plan_global_replacements(requests)

    assert len(plan.assignments) == 6
    assert plan.unresolved_request_ids == ()
    assert len({(item.provider_id, item.replacement_time) for item in plan.assignments}) == 6
