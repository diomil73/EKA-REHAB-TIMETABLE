"""Synthetic exact-membership checks for afternoon shared session review."""
from datetime import time

import pytest

from rehab_excel.afternoon_shared_slot_gate import (
    AfternoonSharedSlotError, SharedSlotApproval, review_afternoon_shared_slots,
)
from rehab_excel.afternoon_slots_source import AfternoonSlot


def slot(patient, *, provider="SYNTH-PT", weekday="ΔΕ", hour=16):
    return AfternoonSlot(patient, "ΦΘ", weekday, time(hour), provider)


def approval(*ids, provider="SYNTH-PT", weekday="ΔΕ", hour=16):
    return SharedSlotApproval(provider, weekday, time(hour), frozenset(ids))


def test_no_shared_slot_needs_no_approval():
    report = review_afternoon_shared_slots((slot("SYNTH-P1"),))
    assert report.ready
    assert report.shared_slot_count == report.matched_approval_count == 0


def test_same_therapist_two_patients_fail_without_exact_membership():
    with pytest.raises(AfternoonSharedSlotError):
        review_afternoon_shared_slots((slot("SYNTH-P1"), slot("SYNTH-P2")))


def test_exact_membership_approves_only_the_intended_group():
    report = review_afternoon_shared_slots(
        (slot("SYNTH-P1"), slot("SYNTH-P2")),
        (approval("SYNTH-P1", "SYNTH-P2"),),
    )
    assert report.ready
    assert (report.shared_slot_count, report.matched_approval_count) == (1, 1)


@pytest.mark.parametrize("members", [
    ("SYNTH-P1",), ("SYNTH-P1", "SYNTH-P3"),
    ("SYNTH-P1", "SYNTH-P2", "SYNTH-P3"),
])
def test_missing_or_stale_membership_fails(members):
    rows = (slot("SYNTH-P1"), slot("SYNTH-P2"))
    if len(members) < 2:
        with pytest.raises(AfternoonSharedSlotError):
            approval(*members)
    else:
        with pytest.raises(AfternoonSharedSlotError):
            review_afternoon_shared_slots(rows, (approval(*members),))


def test_duplicate_approval_rejected():
    rows = (slot("SYNTH-P1"), slot("SYNTH-P2"))
    exact = approval("SYNTH-P1", "SYNTH-P2")
    with pytest.raises(AfternoonSharedSlotError):
        review_afternoon_shared_slots(rows, (exact, exact))


def test_stale_approval_rejected_when_group_dissolves():
    with pytest.raises(AfternoonSharedSlotError):
        review_afternoon_shared_slots(
            (slot("SYNTH-P1"),),
            (approval("SYNTH-P1", "SYNTH-P2"),),
        )


def test_different_therapists_do_not_form_shared_group():
    report = review_afternoon_shared_slots(
        (slot("SYNTH-P1"), slot("SYNTH-P2", provider="SYNTH-OTHER"))
    )
    assert report.ready


def test_morning_data_is_not_part_of_afternoon_gate():
    rows = (slot("SYNTH-P1"),)
    assert review_afternoon_shared_slots(rows).ready


def test_shared_same_provider_different_weekday_needs_no_approval():
    rows = (slot("SYNTH-P1"), slot("SYNTH-P2", weekday="ΤΡ"))
    assert review_afternoon_shared_slots(rows).ready


def test_patient_duplicate_afternoon_slot_cannot_be_approved():
    rows = (slot("SYNTH-P1"), slot("SYNTH-P1"))
    with pytest.raises(AfternoonSharedSlotError):
        review_afternoon_shared_slots(rows, ())


def test_incorrect_approval_object_rejected():
    with pytest.raises(AfternoonSharedSlotError):
        review_afternoon_shared_slots((), (object(),))


def test_patient_identifiers_not_disclosed_in_failure_message():
    with pytest.raises(AfternoonSharedSlotError) as failure:
        review_afternoon_shared_slots(
            (slot("SYNTH-P1"), slot("SYNTH-P2"))
        )
    assert "SYNTH-P1" not in str(failure.value)
    assert "SYNTH-P2" not in str(failure.value)
