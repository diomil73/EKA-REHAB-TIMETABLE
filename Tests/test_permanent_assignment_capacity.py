from datetime import time

from rehab_core.models import BaseScheduleEntry
from rehab_core.permanent_assignment import check_permanent_assignment


def e(i, patient, provider, hhmm, days="Καθ/να"):
    hh, mm = map(int, hhmm.split(":"))
    return BaseScheduleEntry(i, patient, "ΦΘ", time(hh, mm), days, provider)


def test_blocks_seventh_distinct_timeslot():
    entries = [e(str(i), f"p{i}", "T", t) for i, t in enumerate(
        ["08:30", "09:15", "10:00", "10:45", "11:30", "12:15"], 1
    )]
    result = check_permanent_assignment(
        provider_id="T", max_daily_timeslots=6, existing_entries=entries,
        proposed_day_pattern="Καθ/να", proposed_time=time(13, 0), patient_id="new"
    )
    assert not result.allowed
    assert len(result.capacity_issues) == 5
    assert all(day.used_timeslots_after == 7 for day in result.days)


def test_allows_sixth_timeslot_when_provider_has_five():
    entries = [e(str(i), f"p{i}", "T", t) for i, t in enumerate(
        ["08:30", "09:15", "10:00", "10:45", "11:30"], 1
    )]
    result = check_permanent_assignment(
        provider_id="T", max_daily_timeslots=6, existing_entries=entries,
        proposed_day_pattern="Καθ/να", proposed_time=time(12, 15), patient_id="new"
    )
    assert result.allowed
    assert all(day.used_timeslots_after == 6 for day in result.days)


def test_permanent_time_change_excludes_source_entry_before_projection():
    entries = [e("moving", "p0", "T", "08:30")] + [
        e(str(i), f"p{i}", "T", t) for i, t in enumerate(
            ["09:15", "10:00", "10:45", "11:30", "12:15"], 1
        )
    ]
    result = check_permanent_assignment(
        provider_id="T", max_daily_timeslots=6, existing_entries=entries,
        proposed_day_pattern="Καθ/να", proposed_time=time(13, 0),
        source_entry_id="moving", patient_id="p0"
    )
    assert result.allowed
    assert all(day.used_timeslots_before == 5 for day in result.days)
    assert all(day.used_timeslots_after == 6 for day in result.days)


def test_complementary_day_patterns_can_share_same_visual_slot():
    entries = [e("existing", "p1", "T", "09:15", "Τρ-Πε")]
    result = check_permanent_assignment(
        provider_id="T", max_daily_timeslots=6, existing_entries=entries,
        proposed_day_pattern="Δε-Τε-Πα", proposed_time=time(9, 15), patient_id="p2"
    )
    assert result.allowed
    assert not result.conflicts


def test_same_day_same_time_different_patient_is_conflict():
    entries = [e("existing", "p1", "T", "09:15", "Καθ/να")]
    result = check_permanent_assignment(
        provider_id="T", max_daily_timeslots=6, existing_entries=entries,
        proposed_day_pattern="Δε-Τε-Πα", proposed_time=time(9, 15), patient_id="p2"
    )
    assert not result.allowed
    assert len(result.conflicts) == 3
