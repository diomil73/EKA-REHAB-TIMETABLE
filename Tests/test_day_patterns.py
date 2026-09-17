import pytest

from rehab_core.day_patterns import (
    RehabWeekday,
    parse_day_pattern,
    patterns_are_complementary,
    patterns_overlap,
)


def test_parses_real_three_day_pattern():
    assert parse_day_pattern("Δ-Τρ-Πε") == frozenset(
        {
            RehabWeekday.MONDAY,
            RehabWeekday.TUESDAY,
            RehabWeekday.THURSDAY,
        }
    )


def test_parses_real_four_day_pattern_with_short_friday():
    assert parse_day_pattern("Δ-Τρ-Πε-Π") == frozenset(
        {
            RehabWeekday.MONDAY,
            RehabWeekday.TUESDAY,
            RehabWeekday.THURSDAY,
            RehabWeekday.FRIDAY,
        }
    )


@pytest.mark.parametrize(
    ("pattern", "expected"),
    [
        ("Δε", {RehabWeekday.MONDAY}),
        ("Δε-Πα", {RehabWeekday.MONDAY, RehabWeekday.FRIDAY}),
        ("Δε-Πε", {RehabWeekday.MONDAY, RehabWeekday.THURSDAY}),
        ("Δε-Τε", {RehabWeekday.MONDAY, RehabWeekday.WEDNESDAY}),
        (
            "Δε-Τε-Πα",
            {RehabWeekday.MONDAY, RehabWeekday.WEDNESDAY, RehabWeekday.FRIDAY},
        ),
        (
            "Δε-Τε-Πε",
            {RehabWeekday.MONDAY, RehabWeekday.WEDNESDAY, RehabWeekday.THURSDAY},
        ),
        (
            "Δε-Τε-Πε-Πα",
            {
                RehabWeekday.MONDAY,
                RehabWeekday.WEDNESDAY,
                RehabWeekday.THURSDAY,
                RehabWeekday.FRIDAY,
            },
        ),
    ],
)
def test_parses_known_workbook_patterns(pattern, expected):
    assert parse_day_pattern(pattern) == frozenset(expected)


def test_accepts_typographic_dash_and_whitespace():
    assert parse_day_pattern(" Δε – Τε – Πα ") == parse_day_pattern("Δε-Τε-Πα")


def test_overlap_detects_shared_day():
    assert patterns_overlap("Δε-Τε-Πα", "Τρ-Πα")


def test_complementary_patterns_do_not_share_days():
    assert patterns_are_complementary("Δε-Τε-Πα", "Τρ-Πε")


def test_unknown_day_token_is_rejected():
    with pytest.raises(ValueError, match="Unknown day token"):
        parse_day_pattern("Δε-ΧΧ")
