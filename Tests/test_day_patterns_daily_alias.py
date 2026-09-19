from rehab_core.day_patterns import RehabWeekday, parse_day_pattern


def test_kath_na_means_every_weekday():
    assert parse_day_pattern("Καθ/να") == frozenset(RehabWeekday)


def test_kathimerina_means_every_weekday():
    assert parse_day_pattern("Καθημερινά") == frozenset(RehabWeekday)


def test_existing_patterns_still_work():
    assert parse_day_pattern("Δε-Τε-Πα") == frozenset(
        {
            RehabWeekday.MONDAY,
            RehabWeekday.WEDNESDAY,
            RehabWeekday.FRIDAY,
        }
    )
