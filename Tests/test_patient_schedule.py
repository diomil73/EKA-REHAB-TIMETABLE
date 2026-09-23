from datetime import date, time
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_core.day_patterns import RehabWeekday
from rehab_core.models import BaseScheduleEntry, ReplacementProviderKind, Session
from rehab_core.patient_schedule import (
    filter_replacement_options_for_patient_schedule,
    nonprovider_patient_conflicts_on_date,
    recurring_patient_conflict_ids,
)
from rehab_core.permanent_assignment import check_permanent_assignment
from rehab_core.replacement_options import ReplacementProviderOption


MONDAY = date(2026, 9, 21)
T0915 = time(9, 15)
T1000 = time(10, 0)


def entry(entry_id, treatment, start, pattern, therapist=None, patient="P1"):
    return BaseScheduleEntry(
        base_entry_id=entry_id,
        patient_id=patient,
        treatment=treatment,
        start_time=start,
        day_pattern=pattern,
        therapist_id=therapist,
    )


def option(*slots):
    return ReplacementProviderOption(
        provider_id="T2",
        display_name="T2",
        provider_kind=ReplacementProviderKind.THERAPIST,
        active_sessions=0,
        infectious_sessions=0,
        robotic_sessions=0,
        replacement_sessions=0,
        requested_time=T0915,
        recommended_time=slots[0],
        exact_time_available=T0915 in slots,
        available_timeslots=tuple(slots),
        capacity_limit=6,
        capacity_remaining=6,
    )


def target_session():
    return Session(
        session_id="planner:2:ΦΘ@2026-09-21",
        patient_id="P1",
        therapist_id="T1",
        session_date=MONDAY,
        start_time=T0915,
        treatment="ΦΘ",
    )


def test_nonprovider_specialty_blocks_patient_exact_time_but_keeps_alternative():
    entries = [entry("planner:2:Εργο", "Εργο", T0915, "Καθ/να")]
    filtered = filter_replacement_options_for_patient_schedule(
        [option(T0915, T1000)],
        target_session=target_session(),
        base_entries=entries,
    )
    assert len(filtered) == 1
    assert filtered[0].recommended_time == T1000
    assert filtered[0].exact_time_available is False
    assert filtered[0].available_timeslots == (T1000,)


def test_provider_option_disappears_when_all_patient_times_are_used_elsewhere():
    entries = [
        entry("planner:2:Εργο", "Εργο", T0915, "Καθ/να"),
        entry("planner:2:Λογο", "Λογο", T1000, "Καθ/να"),
    ]
    filtered = filter_replacement_options_for_patient_schedule(
        [option(T0915, T1000)],
        target_session=target_session(),
        base_entries=entries,
    )
    assert filtered == []


def test_provider_owned_entry_is_left_to_existing_operational_availability_engine():
    entries = [entry("planner:2:Ρομποτικό", "Ρομποτικό", T0915, "Καθ/να", "T3")]
    conflicts = nonprovider_patient_conflicts_on_date(
        entries=entries,
        patient_id="P1",
        target_date=MONDAY,
        target_time=T0915,
    )
    assert conflicts == ()


def test_recurring_patient_conflicts_are_reported_only_on_overlapping_days():
    entries = [entry("planner:2:Εργο", "Εργο", T0915, "Τρ-Πε")]
    conflicts = recurring_patient_conflict_ids(
        entries=entries,
        patient_id="P1",
        proposed_day_pattern="Καθ/να",
        proposed_time=T0915,
    )
    assert conflicts[RehabWeekday.MONDAY] == ()
    assert conflicts[RehabWeekday.TUESDAY] == ("planner:2:Εργο",)
    assert conflicts[RehabWeekday.WEDNESDAY] == ()
    assert conflicts[RehabWeekday.THURSDAY] == ("planner:2:Εργο",)
    assert conflicts[RehabWeekday.FRIDAY] == ()


def test_permanent_assignment_is_blocked_by_other_specialty_for_same_patient():
    entries = [
        entry("planner:2:Εργο", "Εργο", T0915, "Καθ/να"),
        entry("planner:3:ΦΘ", "ΦΘ", T1000, "Καθ/να", "T2", patient="P2"),
    ]
    check = check_permanent_assignment(
        provider_id="T2",
        max_daily_timeslots=6,
        existing_entries=entries,
        proposed_day_pattern="Καθ/να",
        proposed_time=T0915,
        patient_id="P1",
    )
    assert check.allowed is False
    assert len(check.patient_conflicts) == 5
    assert all(
        day.patient_conflicting_entry_ids == ("planner:2:Εργο",)
        for day in check.days
    )


def test_permanent_change_does_not_conflict_with_its_own_source_entry():
    entries = [entry("planner:2:ΦΘ", "ΦΘ", T0915, "Καθ/να", "T1")]
    check = check_permanent_assignment(
        provider_id="T2",
        max_daily_timeslots=6,
        existing_entries=entries,
        proposed_day_pattern="Καθ/να",
        proposed_time=T0915,
        source_entry_id="planner:2:ΦΘ",
        patient_id="P1",
    )
    assert check.allowed is True
    assert check.patient_conflicts == ()
