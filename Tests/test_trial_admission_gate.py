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
    assert "provider weekday/time double-bookings" in " ".join(result.blockers)


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
