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
    "9:15 ΦΑΝΤΗ",                    # E: day not stated; not infer daily
    "ΔΕ ΤΡ 11.30 SH",                # H: free-text annotation not validated
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


@pytest.mark.parametrize("raw, column, expected", [
    ("9:15 ΦΑΝΤΗ ΔΕ-ΤΡΙ-ΠΕ", "E", {"ΔΕ": "09:15", "ΤΡ": "09:15", "ΠΕ": "09:15"}),
    ("11:30 ΔΕ - ΤE -ΠA ΦΑΝΤΗ", "E", {"ΔΕ": "11:30", "ΤΕ": "11:30", "ΠΑ": "11:30"}),
    ("10:00 ΓΕΩΡΓΟΣΤΑΘΗ - ΠΙΠΕΡΑΚΗ ΤΕ -ΠΑ", "E", {"ΤΕ": "10:00", "ΠΑ": "10:00"}),
    ("10:00 ΛΑΠΠΑ - ΜΠΕΚΟΥ ΔΕ-ΤΕΤ-ΠΑ (Τ)", "E",
     {"ΔΕ": "10:00", "ΤΕ": "10:00", "ΠΑ": "10:00"}),
    ("ΤΕ, ΠΑ 8.30", "H", {"ΤΕ": "08:30", "ΠΑ": "08:30"}),
    ("ΔΕ ΤΡ ΠΕ 10.00", "H", {"ΔΕ": "10:00", "ΤΡ": "10:00", "ΠΕ": "10:00"}),
])
def test_single_time_weekdays_in_any_position_with_optional_e_names(raw, column, expected):
    result = split_day_times(raw, service_column=column)
    assert result.status == "PROPOSED_SPLIT"
    assert dict(result.assignments) == expected


def test_h_does_not_treat_a_therapist_name_as_slot_metadata():
    result = split_day_times("ΤΕ ΠΑ 8.30 ΦΑΝΤΗ", service_column="H")
    assert result.status == "REVIEW"


def test_e_with_no_explicit_weekdays_never_assumes_daily():
    result = split_day_times("8:30 ΦΑΝΤΗ", service_column="E")
    assert result.status == "REVIEW"
