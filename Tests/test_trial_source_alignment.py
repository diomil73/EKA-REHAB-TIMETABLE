"""Synthetic-only read-only provenance tests; never use real patient data."""
from openpyxl import Workbook, load_workbook

from rehab_excel.trial_source_alignment import (
    PROGRAMME_HEADERS,
    PATIENT_HEADERS,
    WEEKLY_HEADERS,
    audit_trial_source_alignment,
    TrialSourceAlignmentError,
)


def _inputs(tmp_path, *, source_name="DEMO PERSON", review_name="DEMO PERSON",
            source_context="WARD A", review_context="WARD A",
            source_cell="C2 / D2 / E2", specialty="ΦΘ",
            schedule_row=2, include_context=True, source_time=True):
    source = tmp_path / "programme.xlsx"
    review = tmp_path / "review.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "ΠΡΟΓΡΑΜΜΑ"
    ws.append(["ΘΑΛΑΜΟΣ", None, *PROGRAMME_HEADERS])
    row = [None] * 22
    row[0] = source_context if include_context else None
    row[1] = source_name
    row[2] = 9.5
    row[3] = "ΔΕ"
    row[4] = "DEMO THERAPIST"
    if specialty == "ΑΝΑΚΛΙΝΟΜΕΝΟ":
        row[5] = "10.00"
    if not source_time:
        row[2] = None
    ws.append(row)
    wb.save(source)
    wb.close()

    wb = Workbook()
    patients = wb.active
    patients.title = "ΑΣΘΕΝΕΙΣ"
    patients.append([*PATIENT_HEADERS, "Προτεινόμενο ID", "Κατάσταση πρότασης"])
    patients.append([2, "Α", review_name, review_context if include_context else None,
                     None, "ΕΚΚΡΕΜΕΙ", "SUGGESTED-ONLY", "UNCONFIRMED"])

    weekly = wb.create_sheet("ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ")
    weekly.append(WEEKLY_HEADERS)
    weekly.append(["Α", schedule_row, review_name, specialty, "ΔΕ", "09:30",
                   "DEMO THERAPIST", None, source_cell, "private raw source text",
                   "PROPOSED_EXACT_SOURCE"])
    wb.save(review)
    wb.close()
    return source, review


def test_exact_patient_and_weekly_source_references_align_read_only(tmp_path):
    source, review = _inputs(tmp_path)
    before = source.read_bytes(), review.read_bytes()
    result = audit_trial_source_alignment(programme=source, review=review)
    assert result.aligned
    assert result.patient_rows_checked == 1
    assert result.weekly_rows_checked == 1
    assert result.source_references_checked == 3
    assert result.identity_mismatches == 0
    assert result.source_reference_mismatches == 0
    assert result.inputs_unchanged
    assert (source.read_bytes(), review.read_bytes()) == before


def test_patient_name_mismatch_fails_closed_without_leaking_name(tmp_path):
    source, review = _inputs(tmp_path, review_name="WRONG PRIVATE NAME")
    result = audit_trial_source_alignment(programme=source, review=review)
    assert not result.aligned
    assert result.identity_mismatches == 1
    assert "PATIENT_SOURCE_ROW_OR_CONTEXT_MISMATCH" in result.blockers
    assert "WRONG PRIVATE NAME" not in repr(result)


def test_patient_context_mismatch_fails_closed(tmp_path):
    source, review = _inputs(tmp_path, review_context="OTHER WARD")
    result = audit_trial_source_alignment(programme=source, review=review)
    assert not result.aligned
    assert result.identity_mismatches == 1


def test_missing_source_context_is_counted_not_forward_filled(tmp_path):
    source, review = _inputs(tmp_path, source_context="", review_context="", include_context=False)
    result = audit_trial_source_alignment(programme=source, review=review)
    assert result.aligned
    assert result.missing_patient_context_rows == 1


def test_weekly_row_must_reference_the_same_source_patient(tmp_path):
    source, review = _inputs(tmp_path, schedule_row=3)
    result = audit_trial_source_alignment(programme=source, review=review)
    assert not result.aligned
    assert result.schedule_identity_mismatches == 1
    assert "WEEKLY_SOURCE_IDENTITY_MISMATCH" in result.blockers


def test_weekly_source_cells_must_match_treatment_columns_and_row(tmp_path):
    source, review = _inputs(tmp_path, source_cell="C3 / D2 / E2")
    result = audit_trial_source_alignment(programme=source, review=review)
    assert not result.aligned
    assert result.source_reference_mismatches == 1
    assert "WEEKLY_SOURCE_CELL_REFERENCE_MISMATCH" in result.blockers


def test_unsupported_treatment_is_held(tmp_path):
    source, review = _inputs(tmp_path, specialty="ΑΠΟΓΕΥΜΑ", source_cell="U2")
    result = audit_trial_source_alignment(programme=source, review=review)
    assert not result.aligned
    assert result.unsupported_specialties == 1
    assert "UNSUPPORTED_WEEKLY_SPECIALTY" in result.blockers


def test_reclining_source_reference_is_time_only(tmp_path):
    source, review = _inputs(tmp_path, specialty="ΑΝΑΚΛΙΝΟΜΕΝΟ", source_cell="F2")
    result = audit_trial_source_alignment(programme=source, review=review)
    assert result.aligned
    assert result.source_references_checked == 1


def test_missing_source_time_cell_fails_closed(tmp_path):
    source, review = _inputs(tmp_path, source_time=False)
    result = audit_trial_source_alignment(programme=source, review=review)
    assert not result.aligned
    assert result.source_reference_mismatches == 1


def test_wrong_programme_headers_stop_before_alignment(tmp_path):
    source, review = _inputs(tmp_path)
    wb = load_workbook(source)
    wb["ΠΡΟΓΡΑΜΜΑ"]["V1"] = "UNKNOWN"
    wb.save(source)
    wb.close()
    try:
        audit_trial_source_alignment(programme=source, review=review)
    except TrialSourceAlignmentError as exc:
        assert "headers differ" in str(exc)
    else:
        raise AssertionError("Expected source header safety stop")
