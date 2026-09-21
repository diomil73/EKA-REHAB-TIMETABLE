from datetime import time
import sys
from pathlib import Path

PYTHON_DIR = Path(__file__).resolve().parents[1] / "Python"
if str(PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(PYTHON_DIR))

from rehab_core.models import BaseScheduleEntry
from rehab_core.permanent_rebalance import (
    find_permanent_rebalance_options,
    provider_weekly_capacity,
)


def entry(i, patient, provider, hhmm, days="Καθ/να"):
    h, m = map(int, hhmm.split(":"))
    return BaseScheduleEntry(
        base_entry_id=i,
        patient_id=patient,
        treatment="ΦΘ",
        start_time=time(h, m),
        day_pattern=days,
        therapist_id=provider,
    )


def test_weekly_capacity_counts_distinct_timeslots():
    entries = [
        entry("a", "p1", "A", "08:30"),
        entry("b", "p2", "A", "08:30", "Τρ-Πε"),
        entry("c", "p3", "A", "09:15"),
    ]
    result = provider_weekly_capacity(
        provider_id="A", existing_entries=entries, max_daily_timeslots=6
    )
    assert [x.used_timeslots for x in result] == [2, 2, 2, 2, 2]


def test_move_must_reduce_source_overcapacity():
    source = [
        entry(str(i), f"p{i}", "A", slot)
        for i, slot in enumerate(
            ["08:30", "09:15", "10:00", "10:45", "11:30", "12:15", "13:00"],
            start=1,
        )
    ]
    # Destination B has room.
    options = find_permanent_rebalance_options(
        source_provider_id="A",
        existing_entries=source,
        destination_provider_ids=["B"],
        standard_timeslots=[time(8, 30), time(9, 15)],
        provider_capacity_limits={"B": 6},
    )
    assert options
    assert all(o.source_excess_before == 5 for o in options)  # 1 excess x 5 weekdays
    assert all(o.source_excess_after == 0 for o in options)
    assert all(o.fully_resolves_source for o in options)


def test_destination_overcapacity_is_blocked():
    source = [
        entry(str(i), f"p{i}", "A", slot)
        for i, slot in enumerate(
            ["08:30", "09:15", "10:00", "10:45", "11:30", "12:15", "13:00"],
            start=1,
        )
    ]
    destination = [
        entry(f"b{i}", f"q{i}", "B", slot)
        for i, slot in enumerate(
            ["08:30", "09:15", "10:00", "10:45", "11:30", "12:15"],
            start=1,
        )
    ]
    options = find_permanent_rebalance_options(
        source_provider_id="A",
        existing_entries=source + destination,
        destination_provider_ids=["B"],
        standard_timeslots=[time(13, 0)],
        provider_capacity_limits={"B": 6},
    )
    assert options == ()


def test_complementary_source_pair_only_frees_its_own_days():
    entries = [
        entry("a", "p1", "A", "08:30", "Δε-Τε-Πα"),
        entry("b", "p2", "A", "08:30", "Τρ-Πε"),
    ]
    # Add six more daily times so source is over on every day.
    for idx, slot in enumerate(["09:15", "10:00", "10:45", "11:30", "12:15", "13:00"], 1):
        entries.append(entry(f"x{idx}", f"x{idx}", "A", slot))
    options = find_permanent_rebalance_options(
        source_provider_id="A",
        existing_entries=entries,
        destination_provider_ids=["B"],
        standard_timeslots=[time(8, 30)],
        provider_capacity_limits={"B": 6},
    )
    a = next(o for o in options if o.source_entry.base_entry_id == "a")
    assert len(a.source_days_freed) == 3
    assert a.source_excess_after == 2  # Tue/Thu remain over until the partner moves too.
    assert not a.fully_resolves_source
