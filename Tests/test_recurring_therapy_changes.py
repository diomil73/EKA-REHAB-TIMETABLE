"""Permanent weekly changes must preserve old service-day facts."""
from datetime import date, time

import pytest

from rehab_core.models import BaseScheduleEntry
from rehab_core.recurring_therapy_changes import (
    RecurringTherapyChange, RecurringTherapyChangeError,
    materialize_effective_sessions, recurring_entries_on_date,
)

START = date(2026, 11, 1)
CHANGE_DAY = date(2026, 11, 19)


def entries():
    return [
        BaseScheduleEntry("weekly:2", "DEMO-P", "ΦΘ", time(9, 15),
                          "Δε-Πε", therapist_id="OLD-T"),
        BaseScheduleEntry("weekly:3", "DEMO-P", "Εργο", time(11, 30),
                          "Πε"),
    ]


def test_therapist_and_clock_change_affect_only_future_sessions():
    change = RecurringTherapyChange(
        "weekly:2", "DEMO-P", CHANGE_DAY,
        new_therapist_id="NEW-T", new_start_time=time(12, 15),
    )
    old = materialize_effective_sessions(
        entries(), date(2026, 11, 16),
        baseline_date=START, changes=[change],
    )
    new = materialize_effective_sessions(
        entries(), CHANGE_DAY,
        baseline_date=START, changes=[change],
    )
    assert [(x.therapist_id, x.start_time) for x in old] == [
        ("OLD-T", time(9, 15))
    ]
    assert [(x.therapist_id, x.start_time) for x in new] == [
        ("NEW-T", time(12, 15))
    ]
    assert all(x.patient_id == "DEMO-P" for x in [*old, *new])
    assert all(x.session_id.startswith("weekly:2@") for x in [*old, *new])


def test_change_one_weekday_assignment_preserves_other_same_specialty_slot():
    source = [
        BaseScheduleEntry("weekly:2", "DEMO-P", "ΦΘ", time(9, 15),
                          "Δε", therapist_id="OLD-T"),
        BaseScheduleEntry("weekly:4", "DEMO-P", "ΦΘ", time(12, 15),
                          "Πε", therapist_id="OTHER-T"),
    ]
    change = RecurringTherapyChange(
        "weekly:2", "DEMO-P", date(2026, 11, 16),
        new_therapist_id="NEW-T",
    )
    monday = materialize_effective_sessions(
        source, date(2026, 11, 16),
        baseline_date=START, changes=[change],
    )
    thursday = materialize_effective_sessions(
        source, CHANGE_DAY,
        baseline_date=START, changes=[change],
    )
    assert [(x.therapist_id, x.start_time) for x in monday] == [
        ("NEW-T", time(9, 15))
    ]
    assert [(x.therapist_id, x.start_time) for x in thursday] == [
        ("OTHER-T", time(12, 15))
    ]


def test_future_change_does_not_mutate_base_entry():
    source = entries()
    original = source[0]
    result = recurring_entries_on_date(
        source, baseline_date=START, target_date=CHANGE_DAY,
        changes=[RecurringTherapyChange(
            "weekly:2", "DEMO-P", CHANGE_DAY, new_therapist_id="NEW-T"
        )],
    )
    assert original.therapist_id == "OLD-T"
    assert result[0].therapist_id == "NEW-T"
    assert result[1] == source[1]


@pytest.mark.parametrize("change", [
    RecurringTherapyChange("unknown", "DEMO-P", CHANGE_DAY, new_start_time=time(12)),
    RecurringTherapyChange("weekly:2", "OTHER", CHANGE_DAY, new_start_time=time(12)),
    RecurringTherapyChange("weekly:2", "DEMO-P", date(2026, 10, 31),
                           new_start_time=time(12)),
])
def test_mismatched_or_backdated_change_rejected(change):
    with pytest.raises(RecurringTherapyChangeError):
        recurring_entries_on_date(
            entries(), baseline_date=START,
            target_date=CHANGE_DAY, changes=[change],
        )


def test_conflicting_same_day_updates_fail_closed():
    changes = [
        RecurringTherapyChange("weekly:2", "DEMO-P", CHANGE_DAY,
                               new_start_time=time(11)),
        RecurringTherapyChange("weekly:2", "DEMO-P", CHANGE_DAY,
                               new_therapist_id="NEW-T"),
    ]
    with pytest.raises(RecurringTherapyChangeError, match="same date"):
        recurring_entries_on_date(
            entries(), baseline_date=START,
            target_date=CHANGE_DAY, changes=changes,
        )


def test_no_empty_permanent_change():
    with pytest.raises(RecurringTherapyChangeError):
        RecurringTherapyChange("weekly:2", "DEMO-P", CHANGE_DAY)
