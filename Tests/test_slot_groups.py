from datetime import date, time

import pytest

from rehab_core.models import BaseScheduleEntry
from rehab_core.slot_groups import build_slot_groups, change_patient_assignment


def entry(
    entry_id: str,
    patient_id: str,
    days: str,
    *,
    therapist: str = "ΘΕΡΑΠΕΥΤΗΣ Α",
    start: time = time(12, 15),
) -> BaseScheduleEntry:
    return BaseScheduleEntry(
        base_entry_id=entry_id,
        patient_id=patient_id,
        treatment="ΦΘ",
        start_time=start,
        day_pattern=days,
        therapist_id=therapist,
    )


def test_complementary_patient_assignments_form_one_derived_pair():
    groups = build_slot_groups(
        [
            entry("a", "P1", "Δε-Τε-Πα"),
            entry("b", "P2", "Τρ-Πε"),
        ]
    )

    assert len(groups) == 1
    assert groups[0].is_pair
    assert groups[0].is_conflict_free
    assert {member.patient_id for member in groups[0].members} == {"P1", "P2"}


def test_daily_view_selects_only_the_member_who_works_that_day():
    group = build_slot_groups(
        [
            entry("a", "P1", "Δε-Τε-Πα"),
            entry("b", "P2", "Τρ-Πε"),
        ]
    )[0]

    monday = group.active_members_on(date(2026, 9, 21))
    tuesday = group.active_members_on(date(2026, 9, 22))

    assert [member.patient_id for member in monday] == ["P1"]
    assert [member.patient_id for member in tuesday] == ["P2"]


def test_same_day_overlap_is_flagged_instead_of_hidden():
    group = build_slot_groups(
        [
            entry("a", "P1", "Δε-Τε-Πα"),
            entry("b", "P2", "Δε-Πε"),
        ]
    )[0]

    assert not group.is_conflict_free
    assert not group.is_pair
    assert len(group.overlaps) == 1


def test_changing_one_member_therapist_dissolves_pair_automatically():
    original = [
        entry("a", "P1", "Δε-Τε-Πα"),
        entry("b", "P2", "Τρ-Πε"),
    ]

    updated, groups = change_patient_assignment(
        original,
        assignment_id="a",
        therapist_id="ΘΕΡΑΠΕΥΤΗΣ Β",
    )

    assert len(groups) == 2
    assert {group.therapist_id for group in groups} == {
        "ΘΕΡΑΠΕΥΤΗΣ Α",
        "ΘΕΡΑΠΕΥΤΗΣ Β",
    }
    assert original[0].therapist_id == "ΘΕΡΑΠΕΥΤΗΣ Α"  # immutable source input
    assert next(item for item in updated if item.base_entry_id == "a").therapist_id == "ΘΕΡΑΠΕΥΤΗΣ Β"


def test_changing_one_member_time_regroups_without_touching_other_patient():
    original = [
        entry("a", "P1", "Δε-Τε-Πα"),
        entry("b", "P2", "Τρ-Πε"),
    ]

    updated, groups = change_patient_assignment(
        original,
        assignment_id="a",
        start_time=time(13, 0),
    )

    assert len(groups) == 2
    p1 = next(item for item in updated if item.patient_id == "P1")
    p2 = next(item for item in updated if item.patient_id == "P2")
    assert p1.start_time == time(13, 0)
    assert p2.start_time == time(12, 15)


def test_change_day_pattern_can_reform_pair_and_is_validated():
    original = [
        entry("a", "P1", "Δε-Τε-Πα"),
        entry("b", "P2", "Δε-Πε"),
    ]
    assert not build_slot_groups(original)[0].is_conflict_free

    _, groups = change_patient_assignment(
        original,
        assignment_id="b",
        day_pattern="Τρ-Πε",
    )
    assert groups[0].is_pair

    with pytest.raises(ValueError):
        change_patient_assignment(
            original,
            assignment_id="b",
            day_pattern="ΑΓΝΩΣΤΗ-ΜΕΡΑ",
        )


def test_unknown_assignment_is_not_silently_ignored():
    with pytest.raises(KeyError):
        change_patient_assignment(
            [entry("a", "P1", "Δε")],
            assignment_id="missing",
            therapist_id="ΘΕΡΑΠΕΥΤΗΣ Β",
        )


def test_same_patient_complementary_assignments_are_not_a_pair():
    group = build_slot_groups(
        [
            entry("a", "P1", "Δε-Τε-Πα"),
            entry("b", "P1", "Τρ-Πε"),
        ]
    )[0]

    assert group.is_conflict_free
    assert group.is_same_patient_split
    assert not group.is_pair
    assert group.distinct_patient_ids == frozenset({"P1"})
