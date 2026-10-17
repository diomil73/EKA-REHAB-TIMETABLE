from datetime import time

from rehab_core.models import BaseScheduleEntry
from rehab_excel.outpatient_schedule_suggestions import (
    SUGGESTION_MODE_LABELS,
    OutpatientScheduleSuggestion,
    SuggestionMode,
    suggest_outpatient_schedule_slots,
)


def entry(therapist: str, at: time, days: str, *, patient: str = "P") -> BaseScheduleEntry:
    return BaseScheduleEntry(
        base_entry_id=f"test:{therapist}:{at}:{days}",
        patient_id=patient,
        treatment="ΦΘ",
        start_time=at,
        day_pattern=days,
        therapist_id=therapist,
        robotic=False,
    )


def suggestions(mode: SuggestionMode, *, entries=()):
    return suggest_outpatient_schedule_slots(
        mode=mode,
        preferred_therapist="T1",
        preferred_time=time(11, 30),
        preferred_day_pattern="Δε-Τε-Πα",
        therapist_names=("T1", "T2", "T3"),
        timeslots=(time(11, 0), time(11, 30), time(12, 0)),
        day_patterns=("Δε-Τε-Πα", "Τρ-Πε", "Δε-Πα"),
        existing_entries=entries,
        limit=20,
    )


def test_dropdown_order_matches_agreed_seven_modes():
    assert SUGGESTION_MODE_LABELS == (
        "Αλλαγή ώρας",
        "Αλλαγή ημερών",
        "Αλλαγή ώρας και ημερών",
        "Αλλαγή θεραπευτή",
        "Αλλαγή θεραπευτή και ώρας",
        "Αλλαγή θεραπευτή και ημερών",
        "Αλλαγή θεραπευτή, ώρας και ημερών",
    )


def test_change_time_keeps_therapist_and_days_and_ranks_nearest_time():
    result = suggestions(SuggestionMode.CHANGE_TIME)
    assert result[0] == OutpatientScheduleSuggestion("T1", time(11, 0), "Δε-Τε-Πα")
    assert result[1] == OutpatientScheduleSuggestion("T1", time(12, 0), "Δε-Τε-Πα")
    assert all(item.therapist_id == "T1" for item in result)
    assert all(item.day_pattern == "Δε-Τε-Πα" for item in result)


def test_change_days_keeps_therapist_and_time():
    result = suggestions(SuggestionMode.CHANGE_DAYS)
    assert result
    assert all(item.therapist_id == "T1" for item in result)
    assert all(item.start_time == time(11, 30) for item in result)
    assert all(item.day_pattern != "Δε-Τε-Πα" for item in result)


def test_change_time_and_days_keeps_only_therapist():
    result = suggestions(SuggestionMode.CHANGE_TIME_AND_DAYS)
    assert result
    assert all(item.therapist_id == "T1" for item in result)
    assert all(item.start_time != time(11, 30) for item in result)
    assert all(item.day_pattern != "Δε-Τε-Πα" for item in result)


def test_change_therapist_keeps_time_and_days():
    result = suggestions(SuggestionMode.CHANGE_THERAPIST)
    assert result[0] == OutpatientScheduleSuggestion("T2", time(11, 30), "Δε-Τε-Πα")
    assert all(item.therapist_id != "T1" for item in result)
    assert all(item.start_time == time(11, 30) for item in result)
    assert all(item.day_pattern == "Δε-Τε-Πα" for item in result)


def test_change_therapist_and_time_keeps_days():
    result = suggestions(SuggestionMode.CHANGE_THERAPIST_AND_TIME)
    assert result
    assert all(item.therapist_id != "T1" for item in result)
    assert all(item.start_time != time(11, 30) for item in result)
    assert all(item.day_pattern == "Δε-Τε-Πα" for item in result)


def test_change_therapist_and_days_keeps_time():
    result = suggestions(SuggestionMode.CHANGE_THERAPIST_AND_DAYS)
    assert result
    assert all(item.therapist_id != "T1" for item in result)
    assert all(item.start_time == time(11, 30) for item in result)
    assert all(item.day_pattern != "Δε-Τε-Πα" for item in result)


def test_change_all_requires_all_three_dimensions_to_change():
    result = suggestions(SuggestionMode.CHANGE_THERAPIST_TIME_AND_DAYS)
    assert result
    assert all(item.therapist_id != "T1" for item in result)
    assert all(item.start_time != time(11, 30) for item in result)
    assert all(item.day_pattern != "Δε-Τε-Πα" for item in result)


def test_busy_provider_slot_is_not_suggested_when_weekdays_overlap():
    result = suggestions(
        SuggestionMode.CHANGE_THERAPIST,
        entries=(entry("T2", time(11, 30), "Καθ/να"),),
    )
    assert OutpatientScheduleSuggestion("T2", time(11, 30), "Δε-Τε-Πα") not in result
    assert OutpatientScheduleSuggestion("T3", time(11, 30), "Δε-Τε-Πα") in result


def test_complementary_weekdays_do_not_block_same_provider_time():
    result = suggestions(
        SuggestionMode.CHANGE_THERAPIST,
        entries=(entry("T2", time(11, 30), "Τρ-Πε"),),
    )
    assert OutpatientScheduleSuggestion("T2", time(11, 30), "Δε-Τε-Πα") in result


def test_limit_is_respected():
    result = suggest_outpatient_schedule_slots(
        mode=SuggestionMode.CHANGE_THERAPIST_TIME_AND_DAYS,
        preferred_therapist="T1",
        preferred_time=time(11, 30),
        preferred_day_pattern="Δε-Τε-Πα",
        therapist_names=("T1", "T2", "T3"),
        timeslots=(time(10, 30), time(11, 0), time(12, 0), time(12, 30)),
        day_patterns=("Δε-Τε-Πα", "Τρ-Πε", "Δε-Πα", "Τρ"),
        existing_entries=(),
        limit=3,
    )
    assert len(result) == 3
