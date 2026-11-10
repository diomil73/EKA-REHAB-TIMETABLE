"""Contract tests for read-only manual column C interpretation."""
import pytest
from tools.parse_real_manual_schedule import parse_physiotherapy_cell, export_review


@pytest.mark.parametrize("raw,initial_time,person,new_time,new_person", [
    ("09.15 ΨΙΑΧΟΥ (ΚΑΛΥΒΙΩΤΗΣ)", "09:15", "ΨΙΑΧΟΥ", "09:15", "ΚΑΛΥΒΙΩΤΗΣ"),
    ("09:15ΨΙΑΧΟΥ(ΚΑΛΥΒΙΩΤΗΣ)", "09:15", "ΨΙΑΧΟΥ", "09:15", "ΚΑΛΥΒΙΩΤΗΣ"),
    ("0915 ΨΙΑΧΟΥ ΚΑΛΥΒΙΩΤΗΣ", "09:15", "ΨΙΑΧΟΥ", "09:15", "ΚΑΛΥΒΙΩΤΗΣ"),
    ("10.45 ΓΑΒΡΑΣ (8.30 ΖΗΜΑΡΗΣ)", "10:45", "ΓΑΒΡΑΣ", "08:30", "ΖΗΜΑΡΗΣ"),
    ("10..45ΓΑΒΡΑΣ(8:30ΖΗΜΑΡΗΣ)", "10:45", "ΓΑΒΡΑΣ", "08:30", "ΖΗΜΑΡΗΣ"),
    ("1045ΓΑΒΡΑΣ830ΖΗΜΑΡΗΣ", "10:45", "ΓΑΒΡΑΣ", "08:30", "ΖΗΜΑΡΗΣ"),
])
def test_tolerant_linked_replacements(raw, initial_time, person, new_time, new_person):
    result = parse_physiotherapy_cell(raw)
    assert (
        result.status, result.initial_time, result.initial_therapist,
        result.replacement_time, result.replacement_therapist,
    ) == ("PROPOSED_REPLACEMENT", initial_time, person, new_time, new_person)


def test_single_time_and_person_are_only_a_proposal():
    result = parse_physiotherapy_cell("10.45 ΣΑΚΚΟΥ")
    assert (result.status, result.initial_time, result.initial_therapist) == (
        "PROPOSED_INITIAL", "10:45", "ΣΑΚΚΟΥ"
    )


@pytest.mark.parametrize("raw", [
    "10.45 ΔΕ-ΤΕ-ΠΑ",
    "11.30 ΡΑΠΤΗ",
    "10.45 ΓΑΒΡΑΣ 9.00",
    "10.45 ΓΑΒΡΑΣ 08.30 ΖΗΜΑΡΗΣ 09.45 ΧΡΗΣΤΟΥ",
])
def test_ambiguous_or_unknown_rows_remain_blocked(raw):
    assert parse_physiotherapy_cell(raw).status == "REVIEW"


def test_output_cannot_overwrite_existing_file(tmp_path):
    target = tmp_path / "review.csv"
    target.write_text("frozen", encoding="utf-8")
    with pytest.raises(FileExistsError):
        export_review(tmp_path / "missing.xlsx", target)
    assert target.read_text(encoding="utf-8") == "frozen"
