"""Regression for concrete source E104/H15/H107 and unsafe ambiguous day-time cells."""
import pytest
from tools.parse_real_manual_day_time_split import split_day_times, export_review


def test_real_e104_day_groups_split_without_duplicate_booking():
    r = split_day_times("11.30(ΔΕ-ΤΕ-ΠΑ)(12.15 ΤΡ-ΠΕ)ΡΑΠΤΗ")
    assert r.status == "PROPOSED_SPLIT"
    assert dict(r.assignments) == {
        "ΔΕ": "11:30", "ΤΕ": "11:30", "ΠΑ": "11:30",
        "ΤΡ": "12:15", "ΠΕ": "12:15",
    }


def test_real_h15_weekday_before_time_and_no_spaces():
    r = split_day_times("ΤΡ,ΠΑ11.30,ΤΕ12.15")
    assert r.status == "PROPOSED_SPLIT"
    assert dict(r.assignments) == {"ΤΡ": "11:30", "ΠΑ": "11:30", "ΤΕ": "12:15"}


def test_real_h107_day_time_mapping():
    r = split_day_times("ΤΡ 10.00, ΠΕ10.45")
    assert r.status == "PROPOSED_SPLIT"
    assert dict(r.assignments) == {"ΤΡ": "10:00", "ΠΕ": "10:45"}


@pytest.mark.parametrize("value", [
    "ΔΕ ΤΡ ΠΕ 11.30, ΠΕ 12.15 SH",  # H31, overlapping day/extra note
    "9:15 ΦΑΝΤΗ ΤΡ-ΠΕ",             # cannot infer association from free text
    "11.30(ΔΕ-ΤΕ-ΠΑ)(12.15 ΤΡ-ΠΕ)ΡΑΠΤΗ 13.00",
    "ΤΡ SH 11.30, ΤΕ 12.15",
])
def test_ambiguous_cells_do_not_create_slots(value):
    assert split_day_times(value).status == "REVIEW"


def test_export_rejects_existing_output_without_touching_source(tmp_path):
    output = tmp_path / "review.csv"
    output.write_text("frozen")
    with pytest.raises(FileExistsError):
        export_review(tmp_path / "missing.xlsx", output)
    assert output.read_text() == "frozen"
