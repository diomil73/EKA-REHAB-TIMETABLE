"""Focused synthetic tests for conservative MASTER/PATIENT_PLANNER mirror."""
from datetime import time

import pytest

from rehab_excel.trial_demo_planner_mirror import (
    PLANNER_TREATMENT_HEADERS, choose_planner_demo_mirrors,
)
from rehab_core.day_patterns import RehabWeekday, parse_day_pattern


def row(patient, treatment, day, clock="09:30", provider="THERAPIST"):
    hours, minutes = (int(part) for part in clock.split(":"))
    return (patient, treatment, day, time(hours, minutes), provider, "C4")


def test_single_weekday_is_visible_in_legacy_planner():
    result = choose_planner_demo_mirrors([row("DEMO-1", "ΦΘ", "ΤΕ")])
    assert len(result) == 1
    assert result[0].day_pattern == "Τε"
    assert result[0].start_time == time(9, 30)
    assert result[0].therapist == "THERAPIST"


def test_five_explicit_weekdays_render_as_daily_not_weekends():
    result = choose_planner_demo_mirrors(
        [row("DEMO-1", "ΦΘ", day) for day in ("ΠΑ", "ΤΕ", "ΔΕ", "ΠΕ", "ΤΡ")]
    )
    assert result[0].day_pattern == "Καθ/να"
    assert parse_day_pattern(result[0].day_pattern) == frozenset(RehabWeekday)


def test_uneven_days_with_single_time_stay_exact_and_sorted():
    result = choose_planner_demo_mirrors(
        [row("DEMO-1", "Εργο", day, provider=None) for day in ("ΠΑ", "ΔΕ", "ΠΕ")]
    )
    assert result[0].day_pattern == "Δε-Πε-Πα"
    assert parse_day_pattern(result[0].day_pattern) == {
        RehabWeekday.MONDAY, RehabWeekday.THURSDAY, RehabWeekday.FRIDAY
    }


def test_different_times_cannot_be_collapsed_into_one_planner_clock():
    result = choose_planner_demo_mirrors([
        row("DEMO-1", "ΦΘ", "ΔΕ", "09:30"),
        row("DEMO-1", "ΦΘ", "ΤΕ", "10:00"),
    ])
    assert result == ()


def test_different_providers_cannot_be_collapsed_or_misrepresented():
    result = choose_planner_demo_mirrors([
        row("DEMO-1", "ΦΘ", "ΔΕ", provider="THERAPIST A"),
        row("DEMO-1", "ΦΘ", "ΤΕ", provider="THERAPIST B"),
    ])
    assert result == ()


def test_multiple_slots_on_same_day_are_not_collapsed():
    result = choose_planner_demo_mirrors([
        row("DEMO-1", "ΦΘ", "ΔΕ", "09:30"),
        row("DEMO-1", "ΦΘ", "ΔΕ", "09:30"),
    ])
    assert result == ()


def test_unsupported_treatment_remains_only_in_granular_view():
    assert choose_planner_demo_mirrors(
        [row("DEMO-1", "Πισίνα", "ΔΕ", "09:00", provider=None)]
    ) == ()


def test_multiple_patients_and_specialties_are_kept_independent():
    result = choose_planner_demo_mirrors([
        row("DEMO-1", "ΦΘ", "ΔΕ", provider="A"),
        row("DEMO-2", "ΦΘ", "ΔΕ", provider="B"),
        row("DEMO-1", "Λογο", "ΤΡ", provider=None),
    ])
    assert {(item.patient_id, item.treatment) for item in result} == {
        ("DEMO-1", "ΦΘ"), ("DEMO-2", "ΦΘ"), ("DEMO-1", "Λογο")
    }


def test_weekend_source_cannot_reach_planner_mirror():
    with pytest.raises(ValueError, match="weekday"):
        choose_planner_demo_mirrors([row("DEMO-1", "ΦΘ", "ΣΑ")])


def test_source_treatment_headers_match_legacy_reader_contract():
    assert PLANNER_TREATMENT_HEADERS["Ψυχ"] == (
        "Ψυχ_Ώρα", "Ψυχ_Ημέρες", "Ψυχ_Ψυχολόγος"
    )
    assert PLANNER_TREATMENT_HEADERS["ΕΦΑ"][2] is None
