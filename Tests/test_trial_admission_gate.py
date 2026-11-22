"""Synthetic fixtures for real-data trial admission. Never expose patient names."""
from datetime import time

from openpyxl import Workbook

from rehab_excel.trial_admission_gate import evaluate_trial_admission


def make_review(tmp_path, *, confirm=True, holds=False, overlaps=False,
                hold_weekly=False, duplicate_id=False):
    wb = Workbook()
    patients = wb.active
    patients.title = "ΑΣΘΕΝΕΙΣ"
    patients.append(["Γραμμή πηγής", "Κλινική", "Ασθενής (πηγή)",
                     "Θάλαμος", "Αντιστοίχιση PatientID", "Επιβεβαίωση"])
    patients.append([4, "Α", "DEMO PERSON", "WARD", "DEMO-1",
                     "CONFIRMED" if confirm else "ΕΚΚΡΕΜΕΙ"])
    if duplicate_id:
        patients.append([5, "Α", "DEMO OTHER", "WARD", "DEMO-1", "CONFIRMED"])
    schedule = wb.create_sheet("ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ")
    schedule.append(["Κλινική", "Γραμμή", "Ασθενής", "Ειδικότητα", "Ημέρα",
                     "Ώρα", "Θεραπευτής ΦΘ/Ρομποτ.", "Προαιρετικό όνομα",
                     "Κελιά πηγής", "Αρχικό κείμενο", "Κατάσταση"])
    schedule.append(["Α", 4, "DEMO PERSON", "ΦΘ", "ΔΕ", "11:30",
                     "DEMO THERAPIST", None, "C4 / D4 / E4", "11.30 | ΔΕ | T",
                     "HOLD_CONFLICT" if hold_weekly else "PROPOSED_EXACT_SOURCE"])
    holds_sheet = wb.create_sheet("ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ")
    holds_sheet.append(["Γραμμή", "Κλινική", "Ασθενής", "Ειδικότητα",
                        "Λόγος", "Αρχικά δεδομένα", "Κελιά"])
    if holds:
        holds_sheet.append([4, "Α", "DEMO PERSON", "ΦΘ", "NO_TIME",
                            "not available", "C4"])
    overlap_sheet = wb.create_sheet("ΕΠΙΚΑΛΥΨΕΙΣ")
    overlap_sheet.append(["Γραμμή", "Κλινική", "Ασθενής", "Ημέρα",
                          "Ώρα", "Ειδικότητες", "Κελιά"])
    if overlaps:
        overlap_sheet.append([4, "Α", "DEMO PERSON", "ΔΕ", "11:30",
                              "ΦΘ / Εργο", "C4 / G4"])
    path = tmp_path / "review.xlsx"
    wb.save(path)
    return path


def test_complete_review_can_be_admitted_without_writing_workbook(tmp_path):
    path = make_review(tmp_path)
    original = path.read_bytes()
    result = evaluate_trial_admission(path)
    assert result.ready
    assert result.candidate_patients == 1
    assert result.verified_patient_ids == 1
    assert result.weekly_proposals == 1
    assert len(result.admitted_entries) == 1
    assert result.admitted_entries[0].start_time == time(11, 30)
    assert result.admitted_entries[0].source_cell == "C4"
    assert path.read_bytes() == original


def test_unconfirmed_ids_block_import_and_do_not_leak_names(tmp_path):
    result = evaluate_trial_admission(make_review(tmp_path, confirm=False))
    assert not result.ready
    assert result.admitted_entries == ()
    assert "patient IDs lack explicit confirmation" in " ".join(result.blockers)
    assert "DEMO PERSON" not in repr(result)


def test_all_unresolved_source_cells_block_import(tmp_path):
    result = evaluate_trial_admission(make_review(tmp_path, holds=True))
    assert not result.ready
    assert result.unresolved_source_rows == 1
    assert result.admitted_entries == ()


def test_overlap_review_rows_block_import(tmp_path):
    result = evaluate_trial_admission(make_review(tmp_path, overlaps=True))
    assert not result.ready
    assert result.overlap_rows == 1
    assert result.admitted_entries == ()


def test_proposed_row_marked_hold_never_passes(tmp_path):
    result = evaluate_trial_admission(make_review(tmp_path, hold_weekly=True))
    assert not result.ready
    assert result.unadmitted_weekly_rows == 1
    assert result.admitted_entries == ()


def test_duplicate_patient_ids_stop_admission(tmp_path):
    result = evaluate_trial_admission(make_review(tmp_path, duplicate_id=True))
    assert not result.ready
    assert result.duplicate_ids == 1
    assert result.admitted_entries == ()



def test_two_patients_same_therapist_weekday_clock_are_blocked(tmp_path):
    path = make_review(tmp_path)
    from openpyxl import load_workbook
    wb = load_workbook(path)
    patients = wb["ΑΣΘΕΝΕΙΣ"]
    patients.append([5, "Α", "DEMO SECOND", "WARD", "DEMO-2", "CONFIRMED"])
    ws = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    ws.append(["Α", 5, "DEMO SECOND", "ΦΘ", "ΔΕ", "11:30",
               "DEMO THERAPIST", None, "C5 / D5 / E5",
               "11.30 | ΔΕ | T", "PROPOSED_EXACT_SOURCE"])
    wb.save(path)
    result = evaluate_trial_admission(path)
    assert not result.ready
    assert result.admitted_entries == ()
    assert "unapproved shared provider weekday/time slots" in " ".join(result.blockers)


def test_same_therapist_different_weekdays_are_allowed(tmp_path):
    path = make_review(tmp_path)
    from openpyxl import load_workbook
    wb = load_workbook(path)
    patients = wb["ΑΣΘΕΝΕΙΣ"]
    patients.append([5, "Α", "DEMO SECOND", "WARD", "DEMO-2", "CONFIRMED"])
    ws = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    ws.append(["Α", 5, "DEMO SECOND", "ΦΘ", "ΠΕ", "11:30",
               "DEMO THERAPIST", None, "C5 / D5 / E5",
               "11.30 | ΠΕ | T", "PROPOSED_EXACT_SOURCE"])
    wb.save(path)
    result = evaluate_trial_admission(path)
    assert result.ready
    assert len(result.admitted_entries) == 2


def test_exact_shared_slot_approval_admits_independent_patient_sessions(tmp_path):
    from openpyxl import load_workbook
    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΑΣΘΕΝΕΙΣ"].append([5, "Α", "DEMO SECOND", "WARD", "DEMO-2", "CONFIRMED"])
    wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"].append(
        ["Α", 5, "DEMO SECOND", "ΦΘ", "ΔΕ", "11:30",
         "DEMO THERAPIST", None, "C5 / D5 / E5",
         "11.30 | ΔΕ | T", "PROPOSED_EXACT_SOURCE"]
    )
    approval = wb.create_sheet("ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS")
    approval.append(["Therapist", "Weekday", "StartTime", "PatientIDs", "Decision"])
    approval.append(["DEMO THERAPIST", "ΔΕ", "11:30", "DEMO-1|DEMO-2",
                     "APPROVED_SHARED"])
    wb.save(path)
    result = evaluate_trial_admission(path)
    assert result.ready
    assert len(result.admitted_entries) == 2
    assert {x.patient_id for x in result.admitted_entries} == {"DEMO-1", "DEMO-2"}


def test_partial_shared_slot_approval_is_not_authoritative(tmp_path):
    from openpyxl import load_workbook
    path = make_review(tmp_path)
    wb = load_workbook(path)
    for line, pid in [(5, "DEMO-2"), (6, "DEMO-3")]:
        wb["ΑΣΘΕΝΕΙΣ"].append([line, "Α", f"DEMO {line}", "WARD", pid, "CONFIRMED"])
        wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"].append(
            ["Α", line, f"DEMO {line}", "ΦΘ", "ΔΕ", "11:30",
             "DEMO THERAPIST", None, f"C{line} / D{line} / E{line}",
             "11.30 | ΔΕ | T", "PROPOSED_EXACT_SOURCE"]
        )
    sheet = wb.create_sheet("ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS")
    sheet.append(["Therapist", "Weekday", "StartTime", "PatientIDs", "Decision"])
    sheet.append(["DEMO THERAPIST", "ΔΕ", "11:30", "DEMO-1|DEMO-2",
                  "APPROVED_SHARED"])
    wb.save(path)
    result = evaluate_trial_admission(path)
    assert not result.ready
    assert result.admitted_entries == ()
    assert "unapproved shared provider" in " ".join(result.blockers)


def test_stale_shared_slot_approval_blocks_otherwise_safe_programme(tmp_path):
    from openpyxl import load_workbook
    path = make_review(tmp_path)
    wb = load_workbook(path)
    sheet = wb.create_sheet("ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS")
    sheet.append(["Therapist", "Weekday", "StartTime", "PatientIDs", "Decision"])
    sheet.append(["DEMO THERAPIST", "ΔΕ", "11:30", "DEMO-1|DEMO-2",
                  "APPROVED_SHARED"])
    wb.save(path)
    result = evaluate_trial_admission(path)
    assert not result.ready
    assert "stale or mismatched" in " ".join(result.blockers)


def test_shared_approval_without_duplicate_member_ids(tmp_path):
    from openpyxl import load_workbook
    path = make_review(tmp_path)
    wb = load_workbook(path)
    sheet = wb.create_sheet("ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS")
    sheet.append(["Therapist", "Weekday", "StartTime", "PatientIDs", "Decision"])
    sheet.append(["DEMO THERAPIST", "ΔΕ", "11:30", "DEMO-1|DEMO-1",
                  "APPROVED_SHARED"])
    wb.save(path)
    result = evaluate_trial_admission(path)
    assert not result.ready
    assert "invalid or duplicated approval" in " ".join(result.blockers)


def test_identical_speech_therapy_proposals_fail_without_overlap_sheet_rows(tmp_path):
    """Same-provenance repeated specialty is held even if overlap tab is empty."""
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    weekly = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    weekly.delete_rows(2)
    fictional = [
        "Α", 4, "DEMO PERSON", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "ΤΕ", "11:30",
        None, None, "J4 / K4 / L4", "FICTIONAL SOURCE",
        "PROPOSED_EXACT_SOURCE",
    ]
    weekly.append(fictional)
    weekly.append(fictional)
    wb.save(path)
    wb.close()

    result = evaluate_trial_admission(path)
    assert result.overlap_rows == 0
    assert result.weekly_proposals == 2
    assert not result.ready
    assert result.admitted_entries == ()
    assert any("duplicate slot detected" in reason for reason in result.blockers)


def test_same_source_speech_duplicate_normalized_clock_forms_fail_without_overlap_tab(tmp_path):
    """Fictional clone: 11:30 and Excel-styled 11.30 are ONE weekday slot."""
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    weekly = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    weekly.delete_rows(2)
    first = [
        "Α", 4, "DEMO PERSON", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "ΤΕ", "11:30",
        None, None, "J4 / K4 / L4", "FICTIONAL SOURCE",
        "PROPOSED_EXACT_SOURCE",
    ]
    second = list(first)
    second[5] = "11.30"
    weekly.append(first)
    weekly.append(second)
    wb.save(path)
    wb.close()
    unchanged = path.read_bytes()

    result = evaluate_trial_admission(path)

    assert result.overlap_rows == 0
    assert result.weekly_proposals == 2
    assert not result.ready
    assert result.admitted_entries == ()
    assert any("duplicate slot detected" in reason for reason in result.blockers)
    assert path.read_bytes() == unchanged


def test_cross_specialty_same_patient_normalized_clock_requires_review_even_without_overlap_tab(
    tmp_path,
):
    """Never rely solely on the manually populated ΕΠΙΚΑΛΥΨΕΙΣ tab."""
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    weekly = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    weekly.delete_rows(2)
    weekly.append([
        "Α", 4, "DEMO PERSON", "ΕΡΓΟΘΕΡΑΠΕΙΑ", "ΔΕ", "11:30",
        None, None, "G4 / H4 / I4", "FICTIONAL",
        "PROPOSED_EXACT_SOURCE",
    ])
    weekly.append([
        "Α", 4, "DEMO PERSON", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "ΔΕ", "11.30",
        None, None, "J4 / K4 / L4", "FICTIONAL",
        "PROPOSED_EXACT_SOURCE",
    ])
    wb.save(path)
    wb.close()

    result = evaluate_trial_admission(path)

    assert result.overlap_rows == 0
    assert result.weekly_proposals == 2
    assert not result.ready
    assert result.admitted_entries == ()
    assert any("duplicate slot detected" in reason for reason in result.blockers)


def test_distinct_patients_same_weekday_clock_with_different_services_not_false_overlap(
    tmp_path,
):
    """Same clock is permissible for distinct patients without shared provider."""
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    patients = wb["ΑΣΘΕΝΕΙΣ"]
    patients.append([5, "Α", "DEMO OTHER", "WARD", "DEMO-2", "CONFIRMED"])
    weekly = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    weekly.delete_rows(2)
    weekly.append([
        "Α", 4, "DEMO PERSON", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "ΤΕ", "11:30",
        None, None, "J4 / K4 / L4", "FICTIONAL",
        "PROPOSED_EXACT_SOURCE",
    ])
    weekly.append([
        "Α", 5, "DEMO OTHER", "ΕΡΓΟΘΕΡΑΠΕΙΑ", "ΤΕ", "11.30",
        None, None, "G5 / H5 / I5", "FICTIONAL",
        "PROPOSED_EXACT_SOURCE",
    ])
    wb.save(path)
    wb.close()

    result = evaluate_trial_admission(path)

    assert result.overlap_rows == 0
    assert result.weekly_proposals == 2
    assert result.verified_patient_ids == 2
    assert result.ready
    assert len(result.admitted_entries) == 2


def test_source_reference_gate_rejects_wrong_columns_rows_missing_time_and_duplicate_cells(
    tmp_path,
):
    """Standalone review approval must never promote an unrelated A:V cell."""
    import pytest
    from openpyxl import load_workbook

    bad_references = (
        "C5 / D5 / E5",  # Different original patient line.
        "G4 / H4 / I4",  # Occupational therapy, not physiotherapy.
        "D4 / E4",       # Missing the actual C time cell.
        "C4 / C4 / E4",  # Repeated source provenance.
        "C4 / D5 / E4",  # One plausible ref points to someone else.
        "C4 / D4 / U4",  # Historical afternoon cannot accompany ΦΘ.
        "E4 / D4",       # Therapist/day details are not a clock.
    )
    for original_cells in bad_references:
        path = make_review(tmp_path)
        wb = load_workbook(path)
        wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]["I2"] = original_cells
        wb.save(path)
        wb.close()
        unchanged = path.read_bytes()
        result = evaluate_trial_admission(path)
        assert not result.ready, original_cells
        assert result.unadmitted_weekly_rows == 1
        assert result.admitted_entries == ()
        assert any("invalid source cell references" in item for item in result.blockers)
        assert path.read_bytes() == unchanged


def test_source_provenance_time_cell_selected_independent_of_ref_order(tmp_path):
    """Original clock C4, not the first arbitrary D/E reference, is canonical."""
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]["I2"] = "E4 / D4 / C4"
    wb.save(path)
    wb.close()
    unchanged = path.read_bytes()
    result = evaluate_trial_admission(path)
    assert result.ready
    assert result.admitted_entries[0].source_cell == "C4"
    assert path.read_bytes() == unchanged


def test_global_source_row_cannot_claim_two_patients_under_different_clinics(tmp_path):
    """One physical A:V source line has only one patient, regardless of clinic."""
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΑΣΘΕΝΕΙΣ"].append([4, "Β", "SECOND FICTIONAL PERSON", "WARD", "DEMO-2", "CONFIRMED"])
    wb.save(path)
    wb.close()
    result = evaluate_trial_admission(path)
    assert not result.ready
    assert result.verified_patient_ids == 2
    assert result.admitted_entries == ()
    assert any("duplicate original source row" in item for item in result.blockers)
    assert "SECOND FICTIONAL PERSON" not in repr(result)


def test_invalid_or_boolean_original_source_row_does_not_admit(tmp_path):
    """Wrong source-row types cannot be inferred into an authoritative identity."""
    from openpyxl import load_workbook

    for source_row in (True, "NOT-A-ROW", 4.5):
        path = make_review(tmp_path)
        wb = load_workbook(path)
        wb["ΑΣΘΕΝΕΙΣ"]["A2"] = source_row
        wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]["B2"] = source_row
        wb.save(path)
        wb.close()
        result = evaluate_trial_admission(path)
        assert not result.ready
        assert result.admitted_entries == ()
        assert any("invalid original source row" in item for item in result.blockers)


def test_formula_only_hold_is_not_silently_lost(tmp_path):
    """A formula with no cached value must still count as an unresolved row."""
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ"]["E2"] = '="FICTIONAL SECRET HOLD"'
    wb.save(path)
    wb.close()
    original = path.read_bytes()

    result = evaluate_trial_admission(path)

    assert result.unresolved_source_rows == 1
    assert not result.ready
    assert result.admitted_entries == ()
    assert any("formula or Excel error" in reason for reason in result.blockers)
    assert "FICTIONAL SECRET HOLD" not in repr(result)
    assert path.read_bytes() == original


def test_formula_only_overlap_is_not_silently_lost(tmp_path):
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΕΠΙΚΑΛΥΨΕΙΣ"]["F2"] = '="FICTIONAL OVERLAP"'
    wb.save(path)
    wb.close()
    original = path.read_bytes()

    result = evaluate_trial_admission(path)

    assert result.overlap_rows == 1
    assert not result.ready
    assert result.admitted_entries == ()
    assert any("formula or Excel error" in reason for reason in result.blockers)
    assert path.read_bytes() == original


def test_formula_generated_confirmation_is_not_explicit_operator_approval(tmp_path):
    from openpyxl import load_workbook

    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΑΣΘΕΝΕΙΣ"]["F2"] = '="CONFIRMED"'
    wb.save(path)
    wb.close()

    result = evaluate_trial_admission(path)

    assert result.verified_patient_ids == 0
    assert result.unadmitted_weekly_rows == 1
    assert not result.ready
    assert result.admitted_entries == ()
    assert any("formula or Excel error" in reason for reason in result.blockers)


def test_unconfirmed_id_with_proposed_weekly_rows_counts_as_unadmitted(tmp_path):
    """A manually filled ID is not enough without the explicit decision."""
    path = make_review(tmp_path, confirm=False)
    result = evaluate_trial_admission(path)
    assert result.verified_patient_ids == 0
    assert result.weekly_proposals == 1
    assert result.unadmitted_weekly_rows == 1
    assert result.admitted_entries == ()
    assert not result.ready


def test_excel_error_in_shared_review_never_passes_as_empty(tmp_path):
    from openpyxl import load_workbook
    from rehab_excel.trial_admission_gate import SHARED_HEADERS, SOURCE_SHARED_SHEET

    path = make_review(tmp_path)
    wb = load_workbook(path)
    sheet = wb.create_sheet(SOURCE_SHARED_SHEET)
    sheet.append(SHARED_HEADERS)
    sheet["E2"] = "#REF!"
    wb.save(path)
    wb.close()

    result = evaluate_trial_admission(path)

    assert not result.ready
    assert result.admitted_entries == ()
    assert any("formula or Excel error" in reason for reason in result.blockers)
    assert "#REF!" not in repr(result)


def test_clean_positive_admission_repr_does_not_expose_private_entries(tmp_path):
    """Approved placements are available privately, not in default diagnostics."""
    result = evaluate_trial_admission(make_review(tmp_path))
    assert result.ready
    assert result.admitted_entries[0].patient_id == "DEMO-1"
    assert "DEMO-1" not in repr(result)
    assert "DEMO THERAPIST" not in repr(result)


def test_recurring_provider_label_variants_require_same_exact_shared_approval(tmp_path):
    from openpyxl import load_workbook
    from rehab_excel.trial_admission_gate import SOURCE_SHARED_SHEET, SHARED_HEADERS

    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΑΣΘΕΝΕΙΣ"].append([5, "Α", "FICTIONAL PEER", "WARD", "DEMO-2", "CONFIRMED"])
    weekly = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    weekly["G2"] = "ΔΟΚΙΜΉ. Α"
    weekly.append([
        "Α", 5, "FICTIONAL PEER", "ΦΘ", "ΔΕ", "11:30",
        "δοκιμη α", None, "C5 / D5 / E5", "FICTIONAL", "PROPOSED_EXACT_SOURCE",
    ])
    wb.save(path)
    wb.close()
    blocked = evaluate_trial_admission(path)
    assert not blocked.ready
    assert any("unapproved shared provider" in x for x in blocked.blockers)
    wb = load_workbook(path)
    approval = wb.create_sheet(SOURCE_SHARED_SHEET)
    approval.append(SHARED_HEADERS)
    approval.append(["ΔοκιμήΑ", "ΔΕ", "11:30", "DEMO-1|DEMO-2", "APPROVED_SHARED"])
    wb.save(path)
    wb.close()
    before = path.read_bytes()
    result = evaluate_trial_admission(path)
    assert result.ready
    assert len(result.admitted_entries) == 2
    assert path.read_bytes() == before


def test_normalized_duplicate_shared_approvals_fail_closed(tmp_path):
    from openpyxl import load_workbook
    from rehab_excel.trial_admission_gate import SOURCE_SHARED_SHEET, SHARED_HEADERS

    path = make_review(tmp_path)
    wb = load_workbook(path)
    wb["ΑΣΘΕΝΕΙΣ"].append([5, "Α", "FICTIONAL PEER", "WARD", "DEMO-2", "CONFIRMED"])
    weekly = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    weekly["G2"] = "ΔΟΚΙΜΉ. Α"
    weekly.append(["Α", 5, "FICTIONAL PEER", "ΦΘ", "ΔΕ", "11:30",
                   "δοκιμη α", None, "C5 / D5 / E5", "FICTIONAL", "PROPOSED_EXACT_SOURCE"])
    approval = wb.create_sheet(SOURCE_SHARED_SHEET)
    approval.append(SHARED_HEADERS)
    for name in ("ΔοκιμήΑ", "δοκιμη. α"):
        approval.append([name, "ΔΕ", "11:30", "DEMO-1|DEMO-2", "APPROVED_SHARED"])
    wb.save(path)
    wb.close()
    result = evaluate_trial_admission(path)
    assert not result.ready
    assert any("invalid or duplicated approval" in x for x in result.blockers)
    assert result.admitted_entries == ()


def _make_source_bound_legacy_review(tmp_path, *, specialty="ΕΡΓΟΘΕΡΑΠΕΙΑ"):
    """Only fictional rows, with the exact existing source-alignment schema."""
    from openpyxl import load_workbook
    from rehab_excel.trial_source_alignment import PROGRAMME_HEADERS, PATIENT_HEADERS

    review = make_review(tmp_path)
    wb = load_workbook(review)
    ws = wb["ΑΣΘΕΝΕΙΣ"]
    for col, header in enumerate(PATIENT_HEADERS, 1):
        ws.cell(1, col, header)
    weekly = wb["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    if specialty == "ΕΡΓΟΘΕΡΑΠΕΙΑ":
        weekly["D2"], weekly["G2"], weekly["H2"], weekly["I2"] = (
            specialty, None, "DEMO STAFF", "G/H/I"
        )
    else:
        weekly["D2"], weekly["G2"], weekly["H2"], weekly["I2"] = (
            specialty, None, "DEMO STAFF", "M/N/O"
        )
    wb.save(review)
    wb.close()

    programme = tmp_path / "manual.xlsx"
    src = Workbook()
    sheet = src.active
    sheet.title = "ΠΡΟΓΡΑΜΜΑ"
    sheet.append(["ΘΑΛΑΜΟΣ", "ΑΣΘΕΝΗΣ", *PROGRAMME_HEADERS])
    sheet.cell(4, 1, "WARD")
    sheet.cell(4, 2, "DEMO PERSON")
    sheet.cell(4, 7 if specialty == "ΕΡΓΟΘΕΡΑΠΕΙΑ" else 13, "11:30")
    src.save(programme)
    src.close()
    return programme, review


def test_original_source_bound_legacy_reference_admission(tmp_path):
    programme, review = _make_source_bound_legacy_review(tmp_path)
    before = programme.read_bytes(), review.read_bytes()
    without_source = evaluate_trial_admission(review)
    assert not without_source.ready
    assert without_source.admitted_entries == ()
    with_source = evaluate_trial_admission(review, programme=programme)
    assert with_source.ready
    assert len(with_source.admitted_entries) == 1
    assert with_source.admitted_entries[0].source_cell == "G4"
    assert (programme.read_bytes(), review.read_bytes()) == before


def test_original_source_bound_legacy_reference_clock_mismatch_fails_closed(tmp_path):
    from openpyxl import load_workbook
    programme, review = _make_source_bound_legacy_review(tmp_path, specialty="ΨΥΧΟΛΟΓΙΑ")
    approved = evaluate_trial_admission(review, programme=programme)
    assert approved.ready
    assert approved.admitted_entries[0].source_cell == "M4"
    source = load_workbook(programme)
    source["ΠΡΟΓΡΑΜΜΑ"]["M4"] = "12:30"
    source.save(programme)
    source.close()
    blocked = evaluate_trial_admission(review, programme=programme)
    assert not blocked.ready
    assert blocked.admitted_entries == ()
    assert "source-aligned" in " ".join(blocked.blockers)


def _fictional_r6_decision_review(tmp_path):
    """A real-R6-shaped decision, with no real names or admission approval."""
    from openpyxl import load_workbook
    from rehab_excel.trial_admission_gate import R6_DECISION_HEADERS

    path = make_review(tmp_path)
    book = load_workbook(path)
    weekly = book["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    weekly["D2"] = "ΕΡΓΟΘΕΡΑΠΕΙΑ"
    weekly["G2"] = None
    weekly["H2"] = "DEMO THERAPIST"
    weekly["I2"] = "G4 / H4 / I4"
    weekly["K2"] = "PROPOSED_OPERATOR_CORRECTION"
    ledger = book.create_sheet("ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ")
    ledger.append(R6_DECISION_HEADERS)
    ledger.append([
        "FICTIONAL OPERATOR CORRECTION", 25, 4, "DEMO PERSON",
        "ΕΡΓΟΘΕΡΑΠΕΙΑ", "ΤΡ", "11:30", None, "DEMO THERAPIST",
        "G4 / H4 / I4", "PROPOSED_EXACT_SOURCE",
        "ΔΕ", "11:30", "DEMO THERAPIST",
        "PROPOSED_OPERATOR_CORRECTION",
    ])
    book.save(path)
    book.close()
    return path


def test_r6_decision_readback_is_exact_but_not_admission(tmp_path):
    path = _fictional_r6_decision_review(tmp_path)
    original = path.read_bytes()
    result = evaluate_trial_admission(path)
    assert not result.ready  # no automatic clinical approval by the ledger
    assert result.admitted_entries == ()
    assert not any("R6 decision" in reason for reason in result.blockers)
    assert path.read_bytes() == original


def test_r6_decision_readback_rejects_changed_after_value(tmp_path):
    from openpyxl import load_workbook
    path = _fictional_r6_decision_review(tmp_path)
    book = load_workbook(path)
    book["ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ"]["M2"] = "12:30"
    book.save(path)
    book.close()
    result = evaluate_trial_admission(path)
    assert not result.ready
    assert result.admitted_entries == ()
    assert any("R6 decision" in reason for reason in result.blockers)


def test_r6_decision_readback_rejects_missing_or_duplicate_ledger_row(tmp_path):
    from openpyxl import load_workbook
    path = _fictional_r6_decision_review(tmp_path)
    book = load_workbook(path)
    ledger = book["ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ"]
    ledger.append([cell.value for cell in ledger[2]])
    book.save(path)
    book.close()
    duplicate = evaluate_trial_admission(path)
    assert any("R6 decision" in item for item in duplicate.blockers)
    book = load_workbook(path)
    book["ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ"].delete_rows(2, 2)
    book.save(path)
    book.close()
    missing = evaluate_trial_admission(path)
    assert any("R6 corrected/combined/held" in item for item in missing.blockers)
