"""Synthetic-only regression for original real-manual F/U/V source triage."""
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.trial_source_alignment import PROGRAMME_HEADERS
from rehab_excel.trial_special_fields_audit import (
    SpecialFieldsAuditError, inspect_real_manual_special_fields,
)
from tools.audit_real_manual_specials import main


def _manual(tmp_path: Path) -> Path:
    path = tmp_path / "PRIVATE_SOURCE.xlsx"
    book = Workbook()
    ws = book.active
    ws.title = "ΠΡΟΓΡΑΜΜΑ"
    for i, header in enumerate(PROGRAMME_HEADERS, start=3):
        ws.cell(1, i).value = header
    ws["A2"] = "DEMO CLINIC"
    ws["B2"] = "DEMO PATIENT A"
    ws["F2"] = "09:30"
    ws["U2"] = "PM ORIGINAL TEXT NOT FOR IMPORT"
    ws["V2"] = "11.15"
    ws["B3"] = "DEMO PATIENT B"
    ws["F3"] = "TILT NOTE / UNCLEAR"
    ws["V3"] = "=TIME(10,30,0)"  # no cached value: must still block
    ws["F4"] = "10:30"  # orphaned special field; not silently ignored
    book.save(path)
    book.close()
    return path


def _edit(path, action):
    book = load_workbook(path)
    action(book["ΠΡΟΓΡΑΜΜΑ"])
    book.save(path)
    book.close()


def test_source_only_audit_counts_all_special_fields_without_writing(tmp_path):
    path = _manual(tmp_path)
    original = path.read_bytes()
    result = inspect_real_manual_special_fields(path)
    assert result.patient_rows == 2
    assert result.reclining_rows == 3
    assert result.reclining_clock_parseable == 2
    assert result.reclining_clock_review == 1
    assert result.afternoon_rows == 1
    assert result.pool_rows == 2
    assert result.pool_clock_parseable == 1
    assert result.pool_clock_review == 1
    assert result.formula_fields == 1
    assert result.orphan_rows == 1
    assert result.unresolved_fields == 6
    assert result.source_unchanged
    assert result.eligible_for_special_field_import is False
    assert path.read_bytes() == original
    assert "DEMO PATIENT" not in repr(result)


def test_no_special_cells_is_clear_inventory_but_not_global_import_permission(tmp_path):
    path = _manual(tmp_path)
    _edit(path, lambda sheet: [
        setattr(sheet[f"{letter}{row}"], "value", None)
        for letter in ("F", "U", "V") for row in (2, 3, 4)
    ])
    result = inspect_real_manual_special_fields(path)
    assert result.unresolved_fields == 0
    assert result.eligible_for_special_field_import is False
    assert result.reclining_rows == result.afternoon_rows == result.pool_rows == 0


def test_pool_and_reclining_readable_clock_never_implies_weekday_or_import(tmp_path):
    path = _manual(tmp_path)
    result = inspect_real_manual_special_fields(path)
    assert result.reclining_clock_parseable == 2
    assert result.pool_clock_parseable == 1
    assert result.unresolved_fields >= 3
    assert result.eligible_for_special_field_import is False


def test_wrong_original_source_headers_fail_closed(tmp_path):
    path = _manual(tmp_path)
    _edit(path, lambda sheet: setattr(sheet["F1"], "value", "WRONG HEADER"))
    with pytest.raises(SpecialFieldsAuditError, match="column contract"):
        inspect_real_manual_special_fields(path)


def test_missing_patient_and_formula_cells_are_counted_not_ignored(tmp_path):
    path = _manual(tmp_path)
    result = inspect_real_manual_special_fields(path)
    assert result.orphan_rows == 1
    assert result.formula_fields == 1
    assert result.pool_clock_review == 1


def test_special_source_command_shows_aggregate_counts_only_and_blocks(tmp_path, capsys):
    path = _manual(tmp_path)
    outcome = main(["--programme", str(path)])
    output = capsys.readouterr().out
    assert outcome == 3
    assert "Total special field blockers: 6" in output
    assert "no special treatment admitted" in output
    assert "DEMO PATIENT" not in output
    assert "PM ORIGINAL TEXT" not in output
    assert "DEMO CLINIC" not in output
    assert str(path) not in output


def test_empty_special_command_is_not_full_import_approval(tmp_path, capsys):
    path = _manual(tmp_path)
    _edit(path, lambda sheet: [
        setattr(sheet[f"{letter}{row}"], "value", None)
        for letter in ("F", "U", "V") for row in (2, 3, 4)
    ])
    code = main(["--programme", str(path)])
    msg = capsys.readouterr().out
    assert code == 0
    assert "full source admission still separately required" in msg


def test_missing_source_command_never_prints_path(tmp_path, capsys):
    path = tmp_path / "PRIVATE_HOSPITAL_PATIENT.xlsx"
    code = main(["--programme", str(path)])
    output = capsys.readouterr().out
    assert code == 2
    assert "PRIVATE_HOSPITAL" not in output
    assert str(path) not in output


def test_non_xlsx_source_rejected(tmp_path):
    path = tmp_path / "PRIVATE_SOURCE.xlsm"
    path.write_bytes(b"not the approved manual source")
    with pytest.raises(SpecialFieldsAuditError, match="xlsx"):
        inspect_real_manual_special_fields(path)


def test_source_mutation_is_detected_even_while_read_only(tmp_path, monkeypatch):
    from rehab_excel import trial_special_fields_audit as module

    path = _manual(tmp_path)
    original = module._digest
    count = [0]

    def changing_digest(file):
        count[0] += 1
        if count[0] == 2:
            file.write_bytes(file.read_bytes() + b"changed")
        return original(file)

    monkeypatch.setattr(module, "_digest", changing_digest)
    with pytest.raises(SpecialFieldsAuditError, match="changed during inspection"):
        inspect_real_manual_special_fields(path)


def test_unexpected_workbook_error_is_privacy_sanitized(tmp_path, monkeypatch):
    from rehab_excel import trial_special_fields_audit as module
    path = _manual(tmp_path)

    def unsafe_open(*args, **kwargs):
        raise RuntimeError("DEMO PATIENT PRIVATE SECRET")

    monkeypatch.setattr(module, "load_workbook", unsafe_open)
    with pytest.raises(SpecialFieldsAuditError, match="Cannot inspect") as err:
        inspect_real_manual_special_fields(path)
    assert "PRIVATE SECRET" not in str(err.value)

def _untimed_scope_fixture(tmp_path):
    from hashlib import sha256
    from rehab_excel.trial_special_fields_audit import (
        UNTIMED_SCOPE_SHEET, UNTIMED_SCOPE_HEADERS,
    )
    p = _manual(tmp_path)
    _edit(p, lambda ws: (setattr(ws["G2"], "value", "*"),
                         setattr(ws["N2"], "value", "ΔΕ"),
                         setattr(ws["O2"], "value", "DEMO PSYCH")))
    review = tmp_path / "review.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = UNTIMED_SCOPE_SHEET
    ws.append(UNTIMED_SCOPE_HEADERS)
    digest = sha256(p.read_bytes()).hexdigest()
    ws.append(["G2", digest, "IGNORE_OCCUPATIONAL_MARK"])
    ws.append(["M2", digest, "PSYCHOLOGY_TIME_PENDING"])
    wb.save(review)
    wb.close()
    return p, review


def test_untimed_scope_requires_exact_original_and_is_read_only(tmp_path):
    from rehab_excel.trial_special_fields_audit import verify_untimed_legacy_scope
    p, review = _untimed_scope_fixture(tmp_path)
    before = p.read_bytes(), review.read_bytes()
    assert verify_untimed_legacy_scope(programme=p, review=review) == frozenset({"G2", "M2"})
    assert (p.read_bytes(), review.read_bytes()) == before


@pytest.mark.parametrize("column,value", [
    ("A2", "C2"), ("A2", "G3"), ("B2", "stale"),
    ("C2", "PSYCHOLOGY_TIME_PENDING"), ("C2", "=1+1"),
    ("D2", "unexpected"),
])
def test_untimed_scope_refuses_inexact_or_unsafe_decisions(tmp_path, column, value):
    from rehab_excel.trial_special_fields_audit import verify_untimed_legacy_scope
    p, r = _untimed_scope_fixture(tmp_path)
    wb = load_workbook(r)
    wb.active[column] = value
    wb.save(r)
    wb.close()
    with pytest.raises(SpecialFieldsAuditError):
        verify_untimed_legacy_scope(programme=p, review=r)


def test_untimed_scope_cannot_acknowledge_real_clock(tmp_path):
    from hashlib import sha256
    from rehab_excel.trial_special_fields_audit import verify_untimed_legacy_scope
    p, r = _untimed_scope_fixture(tmp_path)
    _edit(p, lambda ws: setattr(ws["M2"], "value", "10:00"))
    wb = load_workbook(r)
    for row in (2, 3):
        wb.active.cell(row, 2).value = sha256(p.read_bytes()).hexdigest()
    wb.save(r)
    wb.close()
    with pytest.raises(SpecialFieldsAuditError):
        verify_untimed_legacy_scope(programme=p, review=r)
