"""Permanent weekly changes must preserve old service-day facts."""
from datetime import date, time

import pytest

from rehab_core.models import BaseScheduleEntry
from rehab_core.recurring_therapy_changes import (
    RecurringTherapyChange, RecurringTherapyChangeError,
    materialize_effective_sessions, recurring_entries_on_date,
    validate_effective_weekly_patient_slots,
    prepare_validated_recurring_change,
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



def test_permanent_clock_change_cannot_overlap_other_thursday_therapy():
    source = entries()
    change = RecurringTherapyChange(
        "weekly:2", "DEMO-P", CHANGE_DAY,
        new_start_time=time(11, 30),
    )
    effective = recurring_entries_on_date(
        source, baseline_date=START, target_date=CHANGE_DAY,
        changes=[change],
    )
    with pytest.raises(RecurringTherapyChangeError, match="weekday patient-slot"):
        validate_effective_weekly_patient_slots(effective)


def test_same_therapy_different_weekdays_are_not_patient_collision():
    source = [
        BaseScheduleEntry("weekly:2", "DEMO-P", "ΦΘ", time(9, 15), "Δε", "T1"),
        BaseScheduleEntry("weekly:3", "DEMO-P", "ΦΘ", time(9, 15), "Πε", "T2"),
    ]
    validate_effective_weekly_patient_slots(source)


def test_patient_collision_detected_on_other_weekday_not_just_target_date():
    source = [
        BaseScheduleEntry("weekly:2", "DEMO-P", "ΦΘ", time(9, 15), "Δε-Πε", "T1"),
        BaseScheduleEntry("weekly:3", "DEMO-P", "Εργο", time(9, 15), "Πε"),
    ]
    with pytest.raises(RecurringTherapyChangeError, match="weekday patient-slot"):
        validate_effective_weekly_patient_slots(source)



def test_preflight_accepts_safe_future_therapist_and_clock_move():
    from rehab_core.models import Therapist
    change = RecurringTherapyChange(
        "weekly:2", "DEMO-P", CHANGE_DAY,
        new_therapist_id="NEW-T", new_start_time=time(12, 15),
    )
    preview = prepare_validated_recurring_change(
        entries(), baseline_date=START, change=change,
        therapists=[Therapist("NEW-T", "Demo therapist", max_daily_timeslots=6)],
    )
    assert preview.change is change
    assert preview.effective_entries[0].therapist_id == "NEW-T"
    assert preview.affected_weekdays == ("MONDAY", "THURSDAY")


def test_preflight_rejects_provider_collision_on_thursday_only():
    from rehab_core.models import Therapist
    source = entries() + [
        BaseScheduleEntry("weekly:4", "OTHER-P", "ΦΘ", time(12, 15),
                          "Πε", "NEW-T")
    ]
    change = RecurringTherapyChange(
        "weekly:2", "DEMO-P", CHANGE_DAY,
        new_therapist_id="NEW-T", new_start_time=time(12, 15),
    )
    with pytest.raises(RecurringTherapyChangeError, match="therapist collision"):
        prepare_validated_recurring_change(
            source, baseline_date=START, change=change,
            therapists=[Therapist("NEW-T", "Demo therapist", max_daily_timeslots=6)],
        )


def test_preflight_rejects_therapist_daily_capacity_increase():
    from rehab_core.models import Therapist
    source = entries() + [
        BaseScheduleEntry("weekly:4", "OTHER-P", "ΦΘ", time(8, 15), "Πε", "NEW-T"),
        BaseScheduleEntry("weekly:5", "THIRD-P", "ΦΘ", time(10, 15), "Πε", "NEW-T"),
    ]
    change = RecurringTherapyChange(
        "weekly:2", "DEMO-P", CHANGE_DAY,
        new_therapist_id="NEW-T", new_start_time=time(12, 15),
    )
    with pytest.raises(RecurringTherapyChangeError, match="capacity excess"):
        prepare_validated_recurring_change(
            source, baseline_date=START, change=change,
            therapists=[Therapist("NEW-T", "Demo therapist", max_daily_timeslots=2)],
        )


def test_preflight_rejects_unknown_new_therapist():
    with pytest.raises(RecurringTherapyChangeError, match="unknown therapist"):
        prepare_validated_recurring_change(
            entries(), baseline_date=START,
            change=RecurringTherapyChange(
                "weekly:2", "DEMO-P", CHANGE_DAY, new_therapist_id="UNREGISTERED"
            ),
        )


def test_preflight_detects_patient_collision_on_second_weekday():
    from rehab_core.models import Therapist
    with pytest.raises(RecurringTherapyChangeError, match="weekday patient-slot"):
        prepare_validated_recurring_change(
            entries(), baseline_date=START,
            change=RecurringTherapyChange(
                "weekly:2", "DEMO-P", CHANGE_DAY,
                new_start_time=time(11, 30),
            ),
            therapists=[Therapist("OLD-T", "Demo therapist")],
        )
