from datetime import date, time

import pytest

from rehab_core.global_replacement_integration import (
    expand_replacement_options_for_global_planner,
    plan_replacement_sessions_globally,
)
from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    ReplacementProviderKind,
    Session,
    Therapist,
)
from rehab_core.replacement_options import ReplacementProviderOption


DAY = date(2026, 10, 22)


def test_expands_every_feasible_option_time_and_preserves_provider_kind():
    target = Session("S1", "P1", "T-OLD", DAY, time(9, 0))
    option = ReplacementProviderOption(
        provider_id="ST1",
        display_name="Student",
        provider_kind=ReplacementProviderKind.STUDENT,
        active_sessions=2,
        active_timeslots=2,
        infectious_sessions=0,
        robotic_sessions=0,
        replacement_sessions=0,
        requested_time=time(9, 0),
        recommended_time=time(9, 0),
        exact_time_available=True,
        available_timeslots=(time(9, 0), time(10, 0)),
        capacity_limit=5,
        capacity_remaining=3,
    )

    candidates = expand_replacement_options_for_global_planner(target, [option])

    assert [item.replacement_time for item in candidates] == [time(9, 0), time(10, 0)]
    assert all(item.provider_kind == ReplacementProviderKind.STUDENT for item in candidates)
    assert candidates[0].is_original_time is True
    assert candidates[1].is_original_time is False


def test_real_candidate_pipeline_protects_unique_patient_slot_globally():
    target_a = Session("A", "P-A", "T-A", DAY, time(9, 0))
    target_b = Session("B", "P-B", "T-B", DAY, time(9, 0))
    sessions = (
        target_a,
        target_b,
        Session("B-10", "P-B", "T-X", DAY, time(10, 0)),
        Session("B-11", "P-B", "T-Y", DAY, time(11, 0)),
    )
    absences = (
        DailyAbsence(AbsenceKind.THERAPIST, "T-A", DAY),
        DailyAbsence(AbsenceKind.THERAPIST, "T-B", DAY),
    )

    plan = plan_replacement_sessions_globally(
        target_sessions=(target_a, target_b),
        therapists=(
            Therapist("T-A", "A"),
            Therapist("T-B", "B"),
            Therapist("T-R", "Replacement"),
        ),
        sessions=sessions,
        absences=absences,
        timeslots=(time(9, 0), time(10, 0), time(11, 0)),
    )

    assert plan.unresolved_request_ids == ()
    by_request = {item.request_id: item for item in plan.assignments}
    assert by_request["B"].provider_id == "T-R"
    assert by_request["B"].replacement_time == time(9, 0)
    assert by_request["A"].provider_id == "T-R"
    assert by_request["A"].replacement_time in {time(10, 0), time(11, 0)}


def test_global_integration_requires_one_service_date():
    first = Session("A", "P-A", "T-A", DAY, time(9, 0))
    second = Session("B", "P-B", "T-B", date(2026, 10, 23), time(9, 0))

    with pytest.raises(ValueError, match="one service date"):
        plan_replacement_sessions_globally(
            target_sessions=(first, second),
            therapists=(Therapist("T-R", "Replacement"),),
            sessions=(first, second),
        )
