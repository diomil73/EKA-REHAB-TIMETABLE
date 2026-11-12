"""Same patient, different weekday/time: replacement safety integration."""
from datetime import date, time

import pytest

from rehab_core.models import BaseScheduleEntry, Session, Therapist
from rehab_core.replacement_options import find_replacement_options
from rehab_core.replacements import (
    create_replacement_assignment, find_replacement_candidates,
)

MON = date(2026, 11, 16)
THU = date(2026, 11, 19)
CLOCK_A = time(9, 15)
CLOCK_B = time(12, 15)


def weekly(entry_id, day, clock, treatment, therapist=None):
    return BaseScheduleEntry(
        base_entry_id=entry_id, patient_id="P1",
        treatment=treatment, start_time=clock, day_pattern=day,
        therapist_id=therapist,
    )


def target(day, clock):
    return Session(
        session_id=f"weekly:2@{day.isoformat()}",
        patient_id="P1", therapist_id="T-ORIGINAL",
        session_date=day, start_time=clock, treatment="ΦΘ",
    )


def roster():
    return [Therapist("T-ORIGINAL", "Original"),
            Therapist("T-REPLACEMENT", "Replacement")]


def test_monday_replacement_ignores_thursday_commitment():
    entries = [
        weekly("weekly:2", "Δε", CLOCK_A, "ΦΘ", "T-ORIGINAL"),
        weekly("weekly:3", "Πε", CLOCK_B, "Εργο"),
    ]
    daily = target(MON, CLOCK_A)
    suggestions = find_replacement_options(
        daily, roster(), [daily], timeslots=(CLOCK_A, CLOCK_B),
        base_entries=entries,
    )
    assert len(suggestions) == 1
    assert suggestions[0].available_timeslots == (CLOCK_A, CLOCK_B)
    assert suggestions[0].exact_time_available


def test_thursday_occupational_therapy_blocks_only_thursday_time():
    entries = [
        weekly("weekly:2", "Πε", CLOCK_A, "ΦΘ", "T-ORIGINAL"),
        weekly("weekly:3", "Πε", CLOCK_B, "Εργο"),
    ]
    daily = target(THU, CLOCK_A)
    suggestions = find_replacement_candidates(
        daily, roster(), [daily], timeslots=(CLOCK_A, CLOCK_B),
        base_entries=entries,
    )
    assert len(suggestions) == 1
    assert suggestions[0].available_timeslots == (CLOCK_A,)


def test_final_replacement_write_revalidates_day_specific_occupancy():
    entries = [
        weekly("weekly:2", "Πε", CLOCK_A, "ΦΘ", "T-ORIGINAL"),
        weekly("weekly:3", "Πε", CLOCK_B, "Εργο"),
    ]
    daily = target(THU, CLOCK_A)
    with pytest.raises(ValueError, match="Patient slot blocked"):
        create_replacement_assignment(
            replacement_id="R1",
            target_session=daily,
            replacement_therapist_id="T-REPLACEMENT",
            therapists=roster(),
            sessions=[daily],
            replacement_time=CLOCK_B,
            base_entries=entries,
        )
    # The original physical therapy slot itself remains eligible for reassignment.
    result = create_replacement_assignment(
        replacement_id="R2",
        target_session=daily,
        replacement_therapist_id="T-REPLACEMENT",
        therapists=roster(),
        sessions=[daily],
        replacement_time=CLOCK_A,
        base_entries=entries,
    )
    assert result.replacement_time == CLOCK_A


def test_patient_occupied_by_other_provider_specialty_also_blocks():
    entries = [
        weekly("weekly:2", "Δε", CLOCK_A, "ΦΘ", "T-ORIGINAL"),
        weekly("weekly:3", "Δε", CLOCK_B, "Ρομποτικό", "T-ROBOT"),
    ]
    daily = target(MON, CLOCK_A)
    suggestions = find_replacement_options(
        daily, roster(), [daily], timeslots=(CLOCK_A, CLOCK_B),
        base_entries=entries,
    )
    assert suggestions[0].available_timeslots == (CLOCK_A,)
