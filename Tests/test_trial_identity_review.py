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
               "Θάλαμος/ιατρός (πηγή)", "Αντιστοίχιση PatientID", "Επιβεβαίωση"])
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



def test_review_suggestions_cannot_accidentally_open_admission_gate(tmp_path):
    """A proposed old ID in G cannot count as approval of E/F."""
    from rehab_excel.trial_admission_gate import evaluate_trial_admission
    source, old, output = create_inputs(tmp_path)
    wb = load_workbook(source)
    proposed = wb.create_sheet("ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ")
    proposed.append(["Κλινική", "Γραμμή", "Ασθενής", "Ειδικότητα",
                     "Ημέρα", "Ώρα", "Θεραπευτής ΦΘ/Ρομποτ.",
                     "Προαιρετικό όνομα", "Κελιά πηγής",
                     "Αρχικό κείμενο", "Κατάσταση"])
    proposed.append(["Α", 4, "ΑΣΘΕΝΗΣ ΕΝΑ", "ΦΘ", "ΔΕ", "09:15",
                     "DEMO-T", None, "C4 / D4 / E4",
                     "09.15 | ΔΕ | DEMO-T", "PROPOSED_EXACT_SOURCE"])
    holds = wb.create_sheet("ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ")
    holds.append(["Γραμμή", "Κλινική", "Ασθενής", "Ειδικότητα",
                  "Λόγος", "Αρχικά δεδομένα", "Κελιά"])
    overlap = wb.create_sheet("ΕΠΙΚΑΛΥΨΕΙΣ")
    overlap.append(["Γραμμή", "Κλινική", "Ασθενής", "Ημέρα",
                    "Ώρα", "Ειδικότητες", "Κελιά"])
    wb.save(source)
    wb.close()
    create_local_identity_review(
        staging_path=source, old_demo_path=old, output_path=output,
    )
    result = evaluate_trial_admission(output)
    assert result.verified_patient_ids == 0
    assert result.admitted_entries == ()
    assert not result.ready
    assert any("lack explicit confirmation" in item for item in result.blockers)



def test_inpatient_review_cannot_reuse_an_outpatient_identity(tmp_path):
    """Even a perfectly matching name must not inherit an outpatient ID."""
    source, old, output = create_inputs(tmp_path)
    wb = load_workbook(old)
    ws = wb["PATIENTS"]
    ws.cell(1, 6).value = "PatientType"
    ws.cell(2, 6).value = "Εξωτερικός"
    wb.save(old)
    wb.close()
    counts = create_local_identity_review(
        staging_path=source, old_demo_path=old, output_path=output,
    )
    assert counts["exact_unconfirmed_suggestions"] == 0
    review = load_workbook(output, read_only=True, data_only=True)
    try:
        assert review["ΑΣΘΕΝΕΙΣ"]["G2"].value is None
        assert review["ΑΣΘΕΝΕΙΣ"]["H2"].value == "NO_EXACT_OLD_MATCH"
        assert review["ΑΣΘΕΝΕΙΣ"]["F2"].value == "ΕΚΚΡΕΜΕΙ"
    finally:
        review.close()



def test_concurrent_output_creation_is_never_deleted(tmp_path, monkeypatch):
    """If another process wins the exclusive-create race, it owns that file."""
    from pathlib import Path
    source, old, output = create_inputs(tmp_path)
    original_open = Path.open
    competing_data = b"ANOTHER WRITER OWNS THIS"

    def simulate_race(path, mode="r", *args, **kwargs):
        if Path(path) == output and mode == "xb":
            output.write_bytes(competing_data)
            raise FileExistsError("competing writer created output")
        return original_open(path, mode, *args, **kwargs)

    monkeypatch.setattr(Path, "open", simulate_race)
    with pytest.raises(FileExistsError):
        create_local_identity_review(
            staging_path=source, old_demo_path=old, output_path=output,
        )
    assert output.read_bytes() == competing_data
