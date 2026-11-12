"""Identity-review generation tests use synthetic names and isolated temp files."""
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.trial_identity_review import (
    TrialIdentityReviewError, create_local_identity_review,
)


def create_inputs(tmp_path):
    source = tmp_path / "source.xlsx"
    old = tmp_path / "old.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "ΑΣΘΕΝΕΙΣ"
    ws.append(["Γραμμή πηγής", "Κλινική", "Ασθενής (πηγή)",
               "Θάλαμος", "Αντιστοίχιση PatientID", "Επιβεβαίωση"])
    ws.append([4, "Α", "ΑΣΘΕΝΗΣ ΕΝΑ", "A1", None, "ΕΚΚΡΕΜΕΙ"])
    ws.append([5, "Α", "ΑΣΘΕΝΗΣ ΔΥΟ", "A2", None, "ΕΚΚΡΕΜΕΙ"])
    wb.save(source)
    wb.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "PATIENTS"
    ws.append(["PatientID", "Room", "Patient", "Infectious", "Status"])
    ws.append(["OLD-ID-5", "A1", "ΑΣΘΕΝΗΣ ΕΝΑ", "ΟΧΙ", "ΠΑΡΩΝ"])
    wb.save(old)
    wb.close()
    return source, old, tmp_path / "review.xlsx"


def test_review_only_suggests_and_does_not_confirm(tmp_path):
    source, old, output = create_inputs(tmp_path)
    source_bytes, old_bytes = source.read_bytes(), old.read_bytes()
    report = create_local_identity_review(
        staging_path=source, old_demo_path=old, output_path=output,
    )
    assert report == {
        "source_rows": 2, "exact_unconfirmed_suggestions": 1,
        "requires_individual_review": 1,
    }
    assert source.read_bytes() == source_bytes
    assert old.read_bytes() == old_bytes
    wb = load_workbook(output, read_only=True)
    try:
        ws = wb["ΑΣΘΕΝΕΙΣ"]
        assert ws["G2"].value == "OLD-ID-5"
        assert ws["H2"].value == "EXACT_NAME_CANDIDATE_UNCONFIRMED"
        assert ws["G3"].value is None
        assert ws["H3"].value == "NO_EXACT_OLD_MATCH"
        assert ws["E2"].value is None
        assert ws["F2"].value == "ΕΚΚΡΕΜΕΙ"
    finally:
        wb.close()


def test_existing_output_is_never_overwritten(tmp_path):
    source, old, output = create_inputs(tmp_path)
    output.write_bytes(b"DO NOT OVERWRITE")
    with pytest.raises(TrialIdentityReviewError, match="already exists"):
        create_local_identity_review(staging_path=source, old_demo_path=old, output_path=output)
    assert output.read_bytes() == b"DO NOT OVERWRITE"


def test_existing_suggestion_columns_block_rewriting(tmp_path):
    source, old, output = create_inputs(tmp_path)
    wb = load_workbook(source)
    wb["ΑΣΘΕΝΕΙΣ"]["G1"] = "Existing external annotations"
    wb.save(source)
    with pytest.raises(TrialIdentityReviewError, match="already used"):
        create_local_identity_review(staging_path=source, old_demo_path=old, output_path=output)
    assert not output.exists()


def test_output_cannot_be_input_source(tmp_path):
    source, old, _ = create_inputs(tmp_path)
    with pytest.raises(TrialIdentityReviewError, match="overwrite"):
        create_local_identity_review(staging_path=source, old_demo_path=old, output_path=source)
