"""Fictional, pure assignment membership checks: never a native Excel proof."""
from dataclasses import replace
from datetime import date, time

import pytest

from rehab_excel.daily_input_therapist_assignment_membership import (
    TherapistAssignmentFact, TherapistAssignmentMembershipError,
    verify_therapist_assignment_memberships,
)

DAY = date(2026, 11, 25)


def facts():
    return (
        TherapistAssignmentFact("assignment:FAKE0001", "P1", "ΦΘ", "THERAPIST A", DAY, time(8)),
        TherapistAssignmentFact("assignment:FAKE0002", "P2", "ΦΘ", "THERAPIST A", DAY, time(8)),
        TherapistAssignmentFact("assignment:FAKE0003", "P3", "Εργο", "THERAPIST A", DAY, time(9)),
        TherapistAssignmentFact("assignment:FAKE0004", "P4", "Λογο", "THERAPIST B", DAY, time(8)),
    )


def check(expected=None, observed=None, approvals=None):
    return verify_therapist_assignment_memberships(
        service_date=DAY,
        expected_assignments=facts() if expected is None else expected,
        observed_assignments=facts() if observed is None else observed,
        approved_shared_memberships=(
            (("THERAPIST A", time(8), frozenset({"P1", "P2"})),)
            if approvals is None else approvals
        ),
    )


def test_exact_assignment_ids_patient_members_and_approved_sharing():
    result = check(observed=tuple(reversed(facts())))
    assert result.compared_assignments == 4
    assert result.distinct_provider_clocks == 3
    assert result.approved_shared_clocks == 1


@pytest.mark.parametrize("field,value", [
    ("assignment_id", "assignment:FAKE9999"),
    ("patient_id", "P999"),
    ("treatment", "Πισίνα"),
    ("provider", "THERAPIST B"),
    ("start_time", time(10)),
    ("service_date", date(2026, 11, 26)),
])
def test_cannot_swap_member_identity_even_when_clock_count_matches(field, value):
    observed = list(facts())
    observed[1] = replace(observed[1], **{field: value})
    with pytest.raises(TherapistAssignmentMembershipError):
        check(observed=tuple(observed))


@pytest.mark.parametrize("observed", [
    (),
    facts()[:-1],
    facts() + (facts()[0],),
    (replace(facts()[0], assignment_id="untrusted-row-index"),) + facts()[1:],
    (replace(facts()[0], start_time=time(8, 0, 1)),) + facts()[1:],
])
def test_missing_duplicate_legacy_or_subminute_membership_fails(observed):
    with pytest.raises(TherapistAssignmentMembershipError):
        check(observed=observed)


@pytest.mark.parametrize("approvals", [
    (),
    (("THERAPIST A", time(8), frozenset({"P1", "P3"})),),
    (("THERAPIST A", time(8), frozenset({"P1", "P2", "P3"})),),
    (("THERAPIST A", time(9), frozenset({"P1", "P2"})),),
    (("THERAPIST A", time(8), frozenset({"P1", "P2"})),) * 2,
    (("THERAPIST A", time(8), frozenset({"P1"})),),
    (("THERAPIST A", time(8), frozenset({"P1", "p1"})),),
])
def test_only_exact_independently_approved_shared_group_allowed(approvals):
    with pytest.raises(TherapistAssignmentMembershipError):
        check(approvals=approvals)


def test_occupied_clock_equality_is_not_membership_equality():
    # Four assignments occupy three provider clocks in both versions.
    # Swapping a patient ID at an occupied clock is still a hard failure.
    wrong = list(facts())
    wrong[1] = replace(wrong[1], patient_id="P5")
    with pytest.raises(TherapistAssignmentMembershipError, match="differs"):
        check(observed=tuple(wrong))


def test_no_authorization_or_workbook_io_in_pure_membership_reconciler():
    import inspect
    import rehab_excel.daily_input_therapist_assignment_membership as module
    source = inspect.getsource(module)
    assert "load_workbook" not in source
    assert "DispatchEx" not in source
    assert "DailyInputJournalApplyEnabled = True" not in source
