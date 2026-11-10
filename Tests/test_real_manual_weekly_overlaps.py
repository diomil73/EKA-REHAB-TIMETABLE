"""Do not promote same-patient/weekday E/H candidate collisions into real bookings."""
import pytest
from tools.audit_real_manual_weekly_overlaps import (
    DaySlot, candidates_from_cells, detect_same_patient_weekly_collisions, export_review,
)


def test_one_patient_two_specialties_same_day_and_time_is_review_only():
    slots = [
        DaySlot("E104", "E", 104, "ΤΡ", "12:15"),
        DaySlot("H104", "H", 104, "ΤΡ", "12:15"),
    ]
    flags = detect_same_patient_weekly_collisions(slots)
    assert len(flags) == 1
    assert (flags[0].patient_row, flags[0].weekday, flags[0].time) == (104, "ΤΡ", "12:15")
    assert flags[0].source_cells == "E104;H104"
    assert flags[0].reason.startswith("REVIEW_ONLY")


def test_same_time_different_patient_rows_are_not_a_collision():
    slots = [
        DaySlot("E104", "E", 104, "ΤΡ", "12:15"),
        DaySlot("H105", "H", 105, "ΤΡ", "12:15"),
    ]
    assert detect_same_patient_weekly_collisions(slots) == []


def test_different_weekdays_or_hours_do_not_collide():
    slots = [
        DaySlot("E104", "E", 104, "ΔΕ", "11:30"),
        DaySlot("H104", "H", 104, "ΤΡ", "11:30"),
        DaySlot("H104", "H", 104, "ΔΕ", "12:15"),
    ]
    assert detect_same_patient_weekly_collisions(slots) == []


def test_unreadable_e_h_source_cells_remain_blocked():
    candidates, blocked = candidates_from_cells({
        "E104": "11.30(ΔΕ-ΤΕ-ΠΑ)(12.15 ΤΡ-ΠΕ)ΡΑΠΤΗ",
        "H104": "ΤΡ 12.15",
        "H31": "ΔΕ ΤΡ ΠΕ 11.30, ΠΕ 12.15 SH",
        "C104": "12.15 ΣΑΚΚΟΥ",
    })
    assert len(candidates) == 6
    assert len(blocked) == 1 and blocked[0][0] == "H31"
    flags = detect_same_patient_weekly_collisions(candidates)
    assert len(flags) == 1 and flags[0].source_cells == "E104;H104"


def test_export_never_overwrites_existing_review(tmp_path):
    output = tmp_path / "review.csv"
    output.write_text("preserve", encoding="utf-8")
    with pytest.raises(FileExistsError):
        export_review(tmp_path / "nonexistent.xlsx", output)
    assert output.read_text(encoding="utf-8") == "preserve"


def test_empty_candidate_list_is_safe():
    assert detect_same_patient_weekly_collisions([]) == []
