"""Synthetic-only contract tests for A:V weekday projection. No real records."""
from datetime import time

import pytest

from rehab_excel.trial_source_weekday_projection import (
    TrialSourceWeekdayError, preview_source_weekdays,
)


def source_row(**cells):
    values = [None] * 22
    values[1] = "DEMO PATIENT"
    for column, value in cells.items():
        values[ord(column) - ord("A")] = value
    return values


def preview(**cells):
    return preview_source_weekdays(
        source_row=4,
        columns=source_row(**cells),
        confirmed_patient_id="DEMO-1",
    )


def test_blank_weekdays_with_time_expand_exactly_monday_to_friday():
    result = preview(C="09:30", E="DEMO THERAPIST")
    assert result.ready_for_independent_review
    assert result.blockers == ()
    assert [(x.weekday, x.start_time, x.source_cell) for x in result.entries] == [
        (day, time(9, 30), "C4") for day in ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")
    ]


def test_blank_clock_and_blank_weekdays_generate_no_phantom_sessions():
    result = preview()
    assert result.entries == ()
    assert result.blockers == ()
    assert not result.ready_for_independent_review


def test_explicit_weekdays_are_not_expanded_or_changed():
    result = preview(C="08:15", D="ΤΡ-ΠΕ", E="DEMO THERAPIST")
    assert [x.weekday for x in result.entries] == ["ΤΡ", "ΠΕ"]
    assert all(x.start_time == time(8, 15) for x in result.entries)


def test_multiple_time_groups_retain_separate_weekday_reservations():
    result = preview(
        G="09:00/10:00", H="ΔΕ-ΤΕ/ΤΡ-ΠΕ", I="DEMO OCCUPATIONAL"
    )
    assert result.blockers == ()
    assert [(x.weekday, x.start_time, x.source_cell) for x in result.entries] == [
        ("ΔΕ", time(9), "G4"),
        ("ΤΕ", time(9), "G4"),
        ("ΤΡ", time(10), "G4"),
        ("ΠΕ", time(10), "G4"),
    ]


def test_missing_clock_with_day_or_provider_blocks_instead_of_creating_slot():
    result = preview(D="ΔΕ-ΤΕ", E="DEMO THERAPIST")
    assert result.entries == ()
    assert result.blockers == ("MISSING_CLOCK_C",)


def test_unknown_clock_or_weekend_day_blocks_without_leaking_cell_text():
    invalid_clock = preview(C="UNRECOGNIZED PATIENT NOTE", E="DEMO THERAPIST")
    weekend = preview(C="09:30", D="ΣΑ", E="DEMO THERAPIST")
    assert invalid_clock.blockers == ("UNPARSEABLE_TREATMENT_C",)
    assert weekend.blockers == ("UNPARSEABLE_TREATMENT_C",)
    assert "PATIENT NOTE" not in repr(invalid_clock)
    assert invalid_clock.entries == weekend.entries == ()


def test_one_blocked_stream_blocks_entire_source_row_not_just_that_stream():
    result = preview(C="09:30", E="DEMO THERAPIST", U="PM NOTE")
    assert result.entries == ()
    assert result.blockers == ("AFTERNOON_UNMAPPED",)
    assert "DEMO-1" not in repr(result)
    assert "PM NOTE" not in repr(result)


def test_reclining_and_pool_wait_for_specific_operational_contracts():
    result = preview(F="11:00", V="POOL SOURCE TEXT")
    assert result.entries == ()
    assert result.blockers == (
        "RECLINING_REQUIRES_SEPARATE_30_MIN_CONTRACT",
        "POOL_UNMAPPED",
    )


def test_missing_fth_provider_is_not_assumed():
    result = preview(C="09:30")
    assert result.entries == ()
    assert result.blockers == ("UNPARSEABLE_TREATMENT_C",)


def test_blank_context_is_not_forward_filled():
    result = preview(C="09:30", E="DEMO THERAPIST")
    assert len(result.entries) == 5


@pytest.mark.parametrize("bad_row", [0, 1, True])
def test_invalid_source_coordinates_fail_closed(bad_row):
    with pytest.raises(TrialSourceWeekdayError, match="Invalid source row"):
        preview_source_weekdays(
            source_row=bad_row, columns=source_row(C="09:30"),
            confirmed_patient_id="DEMO-1",
        )


def test_bad_source_shape_or_missing_confirmed_id_fails_closed():
    with pytest.raises(TrialSourceWeekdayError, match="Invalid source row"):
        preview_source_weekdays(
            source_row=4, columns=[None] * 21,
            confirmed_patient_id="DEMO-1",
        )
    with pytest.raises(TrialSourceWeekdayError, match="Invalid source row"):
        preview_source_weekdays(
            source_row=4, columns=source_row(C="09:30"),
            confirmed_patient_id="",
        )

@pytest.mark.parametrize("clock", ["9\u038415", "09\u038415", "09:15", "09.15", "09,15"])
def test_clock_separator_variants_preserve_exact_weekday_and_source(clock):
    result = preview(G=clock, H="ΤΡ-ΠΕ", I="DEMO STAFF")
    assert result.blockers == ()
    assert [(x.weekday, x.start_time, x.source_cell, x.therapist_name)
            for x in result.entries] == [
        (day, time(9, 15), "G4", "DEMO STAFF") for day in ("ΤΡ", "ΠΕ")
    ]


@pytest.mark.parametrize("clock", ["*", "9\u03841", "9\u038460", "24\u038415",
                                  "9\u038415*", "9\u0384\u038415"])
def test_ambiguous_or_invalid_tonos_clock_remains_blocked(clock):
    result = preview(G=clock, H="ΤΡ-ΠΕ", I="DEMO STAFF")
    assert result.entries == ()
    assert result.blockers == ("UNPARSEABLE_TREATMENT_G",)
