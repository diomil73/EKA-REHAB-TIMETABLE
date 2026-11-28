from pathlib import Path

from rehab_excel.h10_k_schema_compare import (
    CANDIDATE_BASELINES,
    find_master_header_row,
    has_vba_project_member,
)


def test_h10_frozen_candidate_manifest_has_three_pinned_hashes_and_sizes():
    assert len(CANDIDATE_BASELINES) == 3
    assert all(len(item["sha256"]) == 64 for item in CANDIDATE_BASELINES.values())
    assert all(item["size"] > 0 for item in CANDIDATE_BASELINES.values())


def test_master_header_row_uses_exact_room_and_patient_pair():
    cells = {
        "B1": "Title",
        "C1": "Report",
        "B4": "Θάλαμος",
        "C4": "Ασθενής",
    }
    assert find_master_header_row(cells) == 4


def test_master_header_row_does_not_guess_when_identity_headers_are_absent():
    assert find_master_header_row({"B1": "Θάλαμος", "C1": "Patient"}) is None


def test_vba_project_member_recognized_in_standard_ooxml_casing():
    assert has_vba_project_member({"[Content_Types].xml", "xl/vbaProject.bin"})
    assert has_vba_project_member({"xl/vbaproject.bin"})


def test_vba_project_member_rejects_missing_or_similarly_named_members():
    assert not has_vba_project_member({"xl/workbook.xml", "xl/vbaProject.bin.bak"})
