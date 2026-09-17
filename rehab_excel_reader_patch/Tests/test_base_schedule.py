from datetime import date, time

from rehab_core.base_schedule import materialize_sessions_for_date, pattern_applies_on_date
from rehab_core.models import BaseScheduleEntry


def test_daily_pattern_applies_on_weekday_not_weekend():
    assert pattern_applies_on_date("Καθ/να", date(2026, 9, 17))
    assert not pattern_applies_on_date("Καθ/να", date(2026, 9, 19))


def test_materializes_only_entries_for_target_day_and_with_provider():
    entries = [
        BaseScheduleEntry(
            base_entry_id="a",
            patient_id="1",
            treatment="ΦΘ",
            start_time=time(12, 15),
            day_pattern="Δε-Τε-Πα",
            therapist_id="T1",
        ),
        BaseScheduleEntry(
            base_entry_id="b",
            patient_id="2",
            treatment="ΦΘ",
            start_time=time(10, 0),
            day_pattern="Τρ-Πε",
            therapist_id="T2",
        ),
        BaseScheduleEntry(
            base_entry_id="c",
            patient_id="3",
            treatment="Εργο",
            start_time=time(11, 30),
            day_pattern="Καθ/να",
            therapist_id=None,
        ),
    ]

    # Thursday 2026-09-17
    sessions = materialize_sessions_for_date(entries, date(2026, 9, 17))

    assert [session.patient_id for session in sessions] == ["2"]
    assert sessions[0].session_id == "b@2026-09-17"
