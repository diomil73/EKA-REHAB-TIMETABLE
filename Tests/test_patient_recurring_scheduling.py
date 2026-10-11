from datetime import date, time
from types import SimpleNamespace

import pytest

from rehab_excel import patient_recurring_scheduling as recurring


ANCHOR = date(2026, 10, 12)  # Monday


def _candidate(therapist_id, name, load, slots):
    return SimpleNamespace(
        therapist_id=therapist_id,
        display_name=name,
        active_sessions=load,
        available_slots=tuple(slots),
    )


def test_recurring_choices_intersect_slots_across_pattern_days(monkeypatch):
    monkeypatch.setattr(
        recurring,
        "read_settings",
        lambda path: SimpleNamespace(day_patterns=("Δε-Τε-Πα",)),
    )
    by_date = {
        date(2026, 10, 12): (
            _candidate("T1", "Γαύρας", 2, (time(9), time(10))),
            _candidate("T2", "Ζήσης", 1, (time(9),)),
        ),
        date(2026, 10, 14): (
            _candidate("T1", "Γαύρας", 3, (time(10), time(11))),
            _candidate("T2", "Ζήσης", 1, (time(9),)),
        ),
        date(2026, 10, 16): (
            _candidate("T1", "Γαύρας", 1, (time(10),)),
            _candidate("T2", "Ζήσης", 2, (time(9), time(10))),
        ),
    }
    monkeypatch.setattr(
        recurring,
        "workbook_therapist_candidates",
        lambda path, **kwargs: by_date[kwargs["target_date"]],
    )

    choices = recurring.recurring_therapist_choices(
        "dummy.xlsm",
        patient_id="P001",
        treatment="ΦΘ",
        day_pattern="Δε-Τε-Πα",
        anchor_date=ANCHOR,
    )

    assert [item.therapist_id for item in choices] == ["T2", "T1"]
    assert choices[0].available_slots == (time(9),)
    assert choices[0].workload_count == 4
    assert choices[0].workload_label == "Ζήσης 4"
    assert choices[1].available_slots == (time(10),)
    assert choices[1].workload_count == 6


def test_recurring_choice_drops_therapist_without_common_slot(monkeypatch):
    monkeypatch.setattr(
        recurring,
        "read_settings",
        lambda path: SimpleNamespace(day_patterns=("Δε-Τε",)),
    )
    monkeypatch.setattr(
        recurring,
        "workbook_therapist_candidates",
        lambda path, **kwargs: (
            _candidate(
                "T1",
                "Γαύρας",
                1,
                (time(9),) if kwargs["target_date"].weekday() == 0 else (time(10),),
            ),
        ),
    )

    choices = recurring.recurring_therapist_choices(
        "dummy.xlsm",
        patient_id="P001",
        treatment="ΦΘ",
        day_pattern="Δε-Τε",
        anchor_date=ANCHOR,
    )

    assert choices == ()


def test_recurring_choice_rejects_unknown_day_pattern(monkeypatch):
    monkeypatch.setattr(
        recurring,
        "read_settings",
        lambda path: SimpleNamespace(day_patterns=("Δε-Τε-Πα",)),
    )

    with pytest.raises(ValueError, match="not defined in SETTINGS"):
        recurring.recurring_therapist_choices(
            "dummy.xlsm",
            patient_id="P001",
            treatment="ΦΘ",
            day_pattern="Τρ-Πε",
            anchor_date=ANCHOR,
        )
