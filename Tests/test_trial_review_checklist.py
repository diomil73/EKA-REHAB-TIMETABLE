"""Fictional-only review inventory tests."""
import pytest
from openpyxl import Workbook, load_workbook
from rehab_excel.trial_source_alignment import PROGRAMME_HEADERS, PATIENT_HEADERS, WEEKLY_HEADERS
from rehab_excel.trial_admission_gate import (
    SOURCE_PATIENT_SHEET, SOURCE_WEEKLY_SHEET, SOURCE_HOLDS_SHEET, SOURCE_OVERLAP_SHEET,
)
from rehab_excel.trial_review_checklist import (
    create_private_review_checklist, TrialReviewChecklistError, TASK_SHEET, PENDING,
)


def _fixture(tmp_path):
    p, r, out = (tmp_path / n for n in ("programme.xlsx", "review.xlsx", "checklist.xlsx"))
    wb = Workbook()
    ws = wb.active
    ws.title = "ΠΡΟΓΡΑΜΜΑ"
    ws.append([None, None, *PROGRAMME_HEADERS])
    ws["B2"] = "PRIVATE FICTIONAL A"
    ws["J2"] = "11:30"
    ws["F2"] = "10:30"
    ws["U2"] = "FICTIONAL HISTORICAL"
    ws["A3"] = "DEMO CONTEXT"
    ws["B3"] = "PRIVATE FICTIONAL B"
    ws["G3"] = "09:30"
    ws["J3"] = "09:30"
    ws["V3"] = "=TIME(9,30,0)"
    wb.save(p)
    wb.close()
    wb = Workbook()
    ws = wb.active
    ws.title = SOURCE_PATIENT_SHEET
    ws.append([*PATIENT_HEADERS, "Suggestion", "Suggestion state"])
    ws.append([2, "A", "PRIVATE FICTIONAL A", None, None, "ΕΚΚΡΕΜΕΙ", "SECRET OLD ID", "UNCONFIRMED"])
    ws.append([3, "A", "PRIVATE FICTIONAL B", "DEMO CONTEXT", "PRIVATE-ID-2", "CONFIRMED"])
    w = wb.create_sheet(SOURCE_WEEKLY_SHEET)
    w.append(WEEKLY_HEADERS)
    for number, name, specialty, clock, refs in (
        (2, "PRIVATE FICTIONAL A", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "11:30", "J2"),
        (2, "PRIVATE FICTIONAL A", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "11.30", "J2"),
        (3, "PRIVATE FICTIONAL B", "ΕΡΓΟΘΕΡΑΠΕΙΑ", "09:30", "G3"),
        (3, "PRIVATE FICTIONAL B", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "09.30", "J3"),
    ):
        w.append(["A", number, name, specialty, "ΔΕ", clock, None, None, refs, "PRIVATE RAW", "PROPOSED_EXACT_SOURCE"])
    hold = wb.create_sheet(SOURCE_HOLDS_SHEET)
    hold.append(["Γραμμή", "Κλινική", "Ασθενής", "Ειδικότητα", "Λόγος", "Αρχικά", "Κελιά"])
    hold["E2"] = '="PRIVATE HOLD"'
    overlap = wb.create_sheet(SOURCE_OVERLAP_SHEET)
    overlap.append(["Γραμμή", "Κλινική", "Ασθενής", "Ημέρα", "Ώρα", "Ειδικότητες", "Κελιά"])
    overlap.append([3, "A", "PRIVATE FICTIONAL B", "ΔΕ", "09:30", "DEMO", "G3 / J3"])
    wb.save(r)
    wb.close()
    return p, r, out


def _run(inputs):
    p, r, out = inputs
    return create_private_review_checklist(programme=p, review=r, output=out)


def test_checklist_collects_existing_independent_and_special_holds_without_mutation(tmp_path):
    p, r, out = _fixture(tmp_path)
    before = p.read_bytes(), r.read_bytes()
    result = _run((p, r, out))
    assert result == {
        "F_RECLINING": 1, "U_AFTERNOON": 1, "V_POOL": 1, "RAW_SOURCE": 1,
        "IDENTITY": 1, "CONTEXT": 1, "HOLD": 1, "REPORTED_OVERLAP": 1,
        "SAME_SPECIALTY": 1, "CROSS_SPECIALTY": 1,
    }
    assert (p.read_bytes(), r.read_bytes()) == before
    wb = load_workbook(out, read_only=True)
    try:
        rows = list(wb[TASK_SHEET].iter_rows(min_row=2, values_only=True))
        assert len(rows) == 10
        assert all(row[4] == PENDING and row[5] is None for row in rows)
        assert ("SAME_SPECIALTY", SOURCE_WEEKLY_SHEET, "2,3", "INDEPENDENT_SLOT_COLLISION", PENDING, None) in rows
        rendered = repr(rows)
        for value in ("PRIVATE FICTIONAL", "SECRET OLD ID", "PRIVATE-ID-2", "11:30", "PRIVATE RAW", "PRIVATE HOLD"):
            assert value not in rendered
    finally:
        wb.close()


def test_existing_output_not_overwritten(tmp_path):
    inputs = _fixture(tmp_path)
    inputs[2].write_bytes(b"existing")
    with pytest.raises(TrialReviewChecklistError):
        _run(inputs)
    assert inputs[2].read_bytes() == b"existing"


def test_source_or_git_destination_rejected(tmp_path):
    p, r, out = _fixture(tmp_path)
    before = r.read_bytes()
    with pytest.raises(TrialReviewChecklistError):
        _run((p, r, r))
    assert r.read_bytes() == before
    git = tmp_path / "repo"
    git.mkdir()
    (git / ".git").mkdir()
    with pytest.raises(TrialReviewChecklistError):
        _run((p, r, git / out.name))


def test_unreviewed_source_treatment_row_inventory(tmp_path):
    p, r, out = _fixture(tmp_path)
    wb = load_workbook(p)
    wb["ΠΡΟΓΡΑΜΜΑ"]["H4"] = "ΠΕ"
    wb.save(p)
    wb.close()
    assert _run((p, r, out))["SOURCE_LINK"] == 1


def test_duplicate_confirmed_patient_ids_are_tasks_not_auto_changed(tmp_path):
    p, r, out = _fixture(tmp_path)
    wb = load_workbook(r)
    ws = wb[SOURCE_PATIENT_SHEET]
    ws["E2"] = "private-id-2"
    ws["F2"] = "CONFIRMED"
    wb.save(r)
    wb.close()
    assert _run((p, r, out))["IDENTITY"] == 1
    wb = load_workbook(out, read_only=True)
    try:
        assert any(row[3] == "DUPLICATE_CONFIRMED_ID" for row in wb[TASK_SHEET].iter_rows(values_only=True))
    finally:
        wb.close()


def test_exclusive_creation_race_preserves_other_writers_file(tmp_path, monkeypatch):
    from pathlib import Path
    p, r, out = _fixture(tmp_path)
    original_open = Path.open
    def race(path, mode="r", *args, **kwargs):
        if path == out and mode == "xb":
            out.write_bytes(b"other writer")
            raise FileExistsError("race")
        return original_open(path, mode, *args, **kwargs)
    monkeypatch.setattr(Path, "open", race)
    with pytest.raises(FileExistsError):
        _run((p, r, out))
    assert out.read_bytes() == b"other writer"


def test_source_change_during_output_stops_and_removes_own_candidate(tmp_path, monkeypatch):
    from rehab_excel import trial_review_checklist as module
    p, r, out = _fixture(tmp_path)
    original_save = module.Workbook.save
    def changed_save(self, target):
        original_save(self, target)
        p.write_bytes(p.read_bytes() + b"changed")
    monkeypatch.setattr(module.Workbook, "save", changed_save)
    with pytest.raises(TrialReviewChecklistError, match="changed"):
        _run((p, r, out))
    assert not out.exists()


def test_cli_prints_only_aggregate_counts(tmp_path, capsys):
    from tools.prepare_trial_review_checklist import main
    p, r, out = _fixture(tmp_path)
    assert main(["--programme", str(p), "--review", str(r), "--output", str(out)]) == 0
    msg = capsys.readouterr().out
    assert "NO APPROVAL OR IMPORT" in msg
    assert "SAME_SPECIALTY: 1" in msg
    assert "PRIVATE FICTIONAL" not in msg
    assert str(p) not in msg
    assert str(out) not in msg



def _fictional_r6_review_checklist(tmp_path, *, unresolved_speech=False):
    """Original A:V and R6 review with only synthetic identity/provider data."""
    from rehab_excel.trial_admission_gate import (
        R6_DECISION_HEADERS, R6_ACTIVE_OVERLAPS_HEADERS,
        R6_COMBINED_DECISION, R6_PENDING_SPEECH_DECISION,
    )
    p, review, out = _fixture(tmp_path)
    book = load_workbook(p)
    original = book["ΠΡΟΓΡΑΜΜΑ"]
    for col, value in (
        ("C3", "09:30"), ("D3", "ΔΕ"), ("E3", "DEMO TEAM"),
        ("R3", "09:30"), ("S3", "ΔΕ"), ("T3", "DEMO TEAM"),
    ):
        original[col] = value
    if unresolved_speech:
        for col, value in (
            ("C2", "11:30"), ("D2", "ΔΕ"), ("E2", "DEMO TEAM"),
            ("J2", "11:30"), ("K2", "ΔΕ"), ("L2", "DEMO SPEECH"),
        ):
            original[col] = value
    book.save(p)
    book.close()
    book = load_workbook(review)
    weekly = book[SOURCE_WEEKLY_SHEET]
    weekly.delete_rows(2, weekly.max_row - 1)
    for specialty, refs in (
        ("ΦΘ", "C3 / D3 / E3"),
        ("ΡΟΜΠΟΤΙΚΗ", "R3 / S3 / T3"),
    ):
        weekly.append([
            "A", 3, "PRIVATE FICTIONAL B", specialty, "ΔΕ", "09:30",
            "DEMO TEAM", None, refs, "FICTIONAL",
            "HOLD_APPROVED_COMBINED",
        ])
    ledger = book.create_sheet("ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ")
    ledger.append(R6_DECISION_HEADERS)
    for index, specialty, refs in (
        (1, "ΦΘ", "C3 / D3 / E3"),
        (2, "ΡΟΜΠΟΤΙΚΗ", "R3 / S3 / T3"),
    ):
        ledger.append([
            "FICTIONAL COMBINED DECISION", index + 20, 3,
            "PRIVATE FICTIONAL B", specialty,
            "ΔΕ", "09:30", "DEMO TEAM", None, refs,
            "HOLD_CONFLICT", "ΔΕ", "09:30", "DEMO TEAM",
            "HOLD_APPROVED_COMBINED",
        ])
    active = book.create_sheet("ΕΝΕΡΓΕΣ_ΕΠΙΚΑΛΥΨΕΙΣ")
    active.append(R6_ACTIVE_OVERLAPS_HEADERS)
    active.append([
        3, "PRIVATE FICTIONAL B", "ΔΕ", "09:30",
        "ΦΘ + ΡΟΜΠΟΤΙΚΗ",
        "C3 / D3 / E3 / R3 / S3 / T3", R6_COMBINED_DECISION,
    ])
    if unresolved_speech:
        weekly.append([
            "A", 2, "PRIVATE FICTIONAL A", "ΦΘ", "ΔΕ", "11:30",
            "DEMO TEAM", None, "C2 / D2 / E2", "FICTIONAL", "HOLD_CONFLICT",
        ])
        weekly.append([
            "A", 2, "PRIVATE FICTIONAL A", "ΛΟΓΟΘΕΡΑΠΕΙΑ", "ΔΕ", "11:30",
            None, "DEMO SPEECH", "J2 / K2 / L2", "FICTIONAL",
            "HOLD_RESCHEDULE",
        ])
        ledger.append([
            "FICTIONAL PENDING SPEECH", 23, 2, "PRIVATE FICTIONAL A",
            "ΛΟΓΟΘΕΡΑΠΕΙΑ", "ΔΕ", "11:30", None, "DEMO SPEECH",
            "J2 / K2 / L2", "HOLD_CONFLICT",
            "ΔΕ", "11:30", "DEMO SPEECH", "HOLD_RESCHEDULE",
        ])
        active.append([
            2, "PRIVATE FICTIONAL A", "ΔΕ", "11:30",
            "ΦΘ + ΛΟΓΟΘΕΡΑΠΕΙΑ",
            "C2 / D2 / E2 / J2 / K2 / L2", R6_PENDING_SPEECH_DECISION,
        ])
    book.save(review)
    book.close()
    return p, review, out


def test_r6_checklist_only_counts_real_pending_active_groups(tmp_path):
    from rehab_excel.trial_review_checklist import TASK_SHEET
    p, review, out = _fictional_r6_review_checklist(
        tmp_path, unresolved_speech=True,
    )
    original = p.read_bytes(), review.read_bytes()
    counts = create_private_review_checklist(
        programme=p, review=review, output=out,
    )
    assert counts.get("ACTIVE_OVERLAP") == 1
    assert counts.get("REPORTED_OVERLAP", 0) == 0
    assert counts.get("CROSS_SPECIALTY", 0) == 0
    assert counts.get("WEEKLY", 0) == 0
    assert counts.get("IDENTITY") == 1  # still unconfirmed patient ID
    assert (p.read_bytes(), review.read_bytes()) == original
    result = load_workbook(out, read_only=True)
    try:
        tasks = list(result[TASK_SHEET].iter_rows(min_row=2, values_only=True))
        assert sum(row[0] == "ACTIVE_OVERLAP" for row in tasks) == 1
        assert all(
            "PRIVATE FICTIONAL" not in str(row) and "DEMO TEAM" not in str(row)
            for row in tasks
        )
    finally:
        result.close()


def test_r6_reviewed_combined_pair_is_not_an_unresolved_task(tmp_path):
    p, review, out = _fictional_r6_review_checklist(tmp_path)
    counts = create_private_review_checklist(
        programme=p, review=review, output=out,
    )
    assert counts.get("ACTIVE_OVERLAP", 0) == 0
    assert counts.get("REPORTED_OVERLAP", 0) == 0
    assert counts.get("CROSS_SPECIALTY", 0) == 0
    assert counts.get("WEEKLY", 0) == 0


def test_r6_checklist_rejects_changed_original_combined_provider(tmp_path):
    p, review, out = _fictional_r6_review_checklist(tmp_path)
    source = load_workbook(p)
    source["ΠΡΟΓΡΑΜΜΑ"]["T3"] = "FICTIONAL CHANGED PROVIDER"
    source.save(p)
    source.close()
    with pytest.raises(TrialReviewChecklistError, match="source proof"):
        create_private_review_checklist(
            programme=p, review=review, output=out,
        )
    assert not out.exists()
