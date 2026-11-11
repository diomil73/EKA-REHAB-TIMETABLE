"""Synthetic regression contract for the authoritative clear-column import rules."""
from datetime import time
import pytest

from rehab_core.programme_mapping import (
    ProgrammeMappingError, parse_clock, parse_weekday_groups, split_weekly_cell,
)


@pytest.mark.parametrize("raw,expected", [
    (0.3854166666666667, time(9, 15)),
    (0.4791666666666667, time(11, 30)),
    (11.3, time(11, 30)),
    (12.15, time(12, 15)),
    (10, time(10, 0)),
    ("11.30", time(11, 30)),
    ("09:15", time(9, 15)),
])
def test_excel_numeric_and_literal_clock_are_equivalent(raw, expected):
    assert parse_clock(raw) == expected


def test_one_clock_two_days_one_therapist():
    slots = split_weekly_cell(
        clock="11.30", weekdays="ΔΕ-ΠΕ", therapist="ΣΑΚΚΟΥ",
        requires_therapist=True,
    )
    assert [(x.weekday, x.start_time.strftime("%H:%M"), x.therapist_name) for x in slots] == [
        ("ΔΕ", "11:30", "ΣΑΚΚΟΥ"), ("ΠΕ", "11:30", "ΣΑΚΚΟΥ")
    ]


def test_ordered_day_clock_provider_slashes():
    slots = split_weekly_cell(
        clock="11.30/12.15", weekdays="ΔΕ/ΠΕ", therapist="ΣΑΚΚΟΥ/ΓΑΒΡΑΣ",
        requires_therapist=True,
    )
    assert [(x.weekday, x.start_time, x.therapist_name) for x in slots] == [
        ("ΔΕ", time(11, 30), "ΣΑΚΚΟΥ"),
        ("ΠΕ", time(12, 15), "ΓΑΒΡΑΣ"),
    ]


def test_day_groups_with_one_time_and_optional_provider():
    slots = split_weekly_cell(
        clock=0.3854166666666667, weekdays="ΔΕ-ΤΡΙ-ΠΕ", therapist="ΦΑΝΤΗ",
    )
    assert len(slots) == 3
    assert all(x.start_time == time(9, 15) for x in slots)


def test_missing_days_daily_only_when_authorized_for_physiotherapy():
    slots = split_weekly_cell(
        clock=12.15, weekdays=None, therapist="ΑΡΓΕΝΤΟΣ",
        requires_therapist=True, implicit_daily=True,
    )
    assert [x.weekday for x in slots] == ["ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ"]
    with pytest.raises(ProgrammeMappingError, match="Missing explicit weekday"):
        split_weekly_cell(clock=12.15, weekdays=None, therapist="ΑΡΓΕΝΤΟΣ")


@pytest.mark.parametrize("kwargs", [
    dict(clock="11.30/12.15", weekdays="ΔΕ-ΠΕ", therapist="ΣΑΚΚΟΥ/ΓΑΒΡΑΣ",
         requires_therapist=True),
    dict(clock="11.30", weekdays="ΔΕ/ΠΕ", therapist="ΣΑΚΚΟΥ/ΓΑΒΡΑΣ/ΨΙΑΧΟΥ",
         requires_therapist=True),
    dict(clock="11.30", weekdays="ΔΕ-ΠΕ", requires_therapist=True),
    dict(clock="11.30", weekdays="ΔΕ-ΔΕ", therapist="ΣΑΚΚΟΥ",
         requires_therapist=True),
    dict(clock="11.30", weekdays="ΠΑΠΑΜΑΝΟΥ", therapist="ΣΑΚΚΟΥ",
         requires_therapist=True),
])
def test_unsafe_misalignment_or_identity_is_blocked(kwargs):
    with pytest.raises(ProgrammeMappingError):
        split_weekly_cell(**kwargs)


def test_raw_and_historical_schedule_are_not_modified():
    source = ["11.30/12.15", "ΔΕ/ΠΕ", "ΣΑΚΚΟΥ/ΓΑΒΡΑΣ"]
    before = source.copy()
    split_weekly_cell(clock=source[0], weekdays=source[1],
                      therapist=source[2], requires_therapist=True)
    assert source == before
