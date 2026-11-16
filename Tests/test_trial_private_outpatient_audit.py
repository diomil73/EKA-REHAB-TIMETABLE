"""Synthetic-only real outpatient audit regression; never uses patient data."""
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.trial_private_outpatient_audit import (
    PrivateOutpatientSourceError, inspect_private_outpatient_source,
)
from tools.audit_real_outpatient_source import main


def _fixture(tmp_path: Path, *, appointments=None) -> Path:
    source = tmp_path / "PRIVATE_OUTPATIENT_SYNTHETIC.xlsx"
    book = Workbook()
    patient = book.active
    patient.title = "PATIENTS"
    patient.append((
        "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης",
        "Κατάσταση", "PatientType", "HospitalMRN",
        "ResponsibleDoctor", "SourceContext",
    ))
    patient.append((
        "DEMO-O1", None, "SYNTHETIC PERSON", "Ο",
        "Ενεργός", "Εξωτερικός", None, None, None,
    ))
    weekly = book.create_sheet("OUTPATIENT_SCHEDULE")
    weekly.append((
        "PatientID", "Ασθενής", "Θεραπεία", "Ώρα", "Ημέρες", "Θεραπευτής",
    ))
    if appointments is None:
        appointments = [
            ("DEMO-O1", "SYNTHETIC PERSON", "ΦΘ", "09:15", "Δε-Τε-Πα", "DEMO-T"),
            ("DEMO-O1", "SYNTHETIC PERSON", "Εργο", "10:15", "Τρ-Πε", None),
        ]
    for record in appointments:
        weekly.append(record)
    book.save(source)
    book.close()
    return source


def _edit(source, change):
    book = load_workbook(source)
    change(book)
    book.save(source)
    book.close()


def test_real_outpatient_audit_reuses_existing_two_sheet_schema(tmp_path):
    source = _fixture(tmp_path)
    original = source.read_bytes()
    report = inspect_private_outpatient_source(source)
    assert report.outpatient_patients == 1
    assert report.recurring_treatments == 2
    assert report.weekday_placements == 5
    assert report.distinct_treatments == 2
    assert report.source_unchanged is True
    assert report.admission_authorized is False
    assert source.read_bytes() == original
    assert "SYNTHETIC PERSON" not in repr(report)


def test_private_command_outputs_aggregates_without_patient_information(tmp_path, capsys):
    source = _fixture(tmp_path)
    assert main(["--source", str(source)]) == 0
    output = capsys.readouterr().out
    assert "outpatient patients: 1" in output
    assert "Monday-Friday placements: 5" in output
    assert "import: PENDING" in output
    assert "SYNTHETIC PERSON" not in output
    assert "DEMO-O1" not in output
    assert str(source) not in output


def test_formula_only_patient_or_schedule_is_not_silently_skipped(tmp_path):
    source = _fixture(tmp_path)
    _edit(source, lambda wb: setattr(
        wb["OUTPATIENT_SCHEDULE"]["D2"], "value", "=TIME(9,15,0)",
    ))
    with pytest.raises(PrivateOutpatientSourceError, match="Formula-only"):
        inspect_private_outpatient_source(source)
    source = _fixture(tmp_path)
    _edit(source, lambda wb: setattr(wb["PATIENTS"]["A2"], "value", '="DEMO-O1"'))
    with pytest.raises(PrivateOutpatientSourceError, match="Formula-only"):
        inspect_private_outpatient_source(source)


@pytest.mark.parametrize("alter", [
    lambda wb: setattr(wb["OUTPATIENT_SCHEDULE"]["A2"], "value", "UNKNOWN"),
    lambda wb: setattr(wb["OUTPATIENT_SCHEDULE"]["B2"], "value", "MISMATCH"),
    lambda wb: setattr(wb["OUTPATIENT_SCHEDULE"]["E2"], "value", "ΣΑ-ΚΥ"),
    lambda wb: setattr(wb["OUTPATIENT_SCHEDULE"]["C2"], "value", "Ρομποτικό"),
    lambda wb: setattr(wb["PATIENTS"]["F2"], "value", "Εσωτερικός"),
    lambda wb: setattr(wb["PATIENTS"]["B2"], "value", "A05"),
    lambda wb: setattr(wb["PATIENTS"]["D2"], "value", "Ν"),
])
def test_existing_production_source_violations_fail_closed(tmp_path, alter):
    source = _fixture(tmp_path)
    _edit(source, alter)
    with pytest.raises(PrivateOutpatientSourceError):
        inspect_private_outpatient_source(source)


def test_missing_physio_provider_is_blocked_before_acceptance(tmp_path):
    source = _fixture(tmp_path)
    _edit(source, lambda wb: setattr(wb["OUTPATIENT_SCHEDULE"]["F2"], "value", None))
    with pytest.raises(PrivateOutpatientSourceError, match="provider"):
        inspect_private_outpatient_source(source)


def test_cross_pattern_duplicates_are_detected_per_weekday(tmp_path):
    source = _fixture(tmp_path)
    _edit(source, lambda wb: wb["OUTPATIENT_SCHEDULE"].append(
        ("DEMO-O1", "SYNTHETIC PERSON", "ΦΘ", "09:15", "Τε-Πε", "OTHER-T"),
    ))
    with pytest.raises(PrivateOutpatientSourceError, match="overlapping duplicates"):
        inspect_private_outpatient_source(source)


def test_distinct_weekday_physio_times_are_valid_and_preserved(tmp_path):
    source = _fixture(tmp_path)
    _edit(source, lambda wb: wb["OUTPATIENT_SCHEDULE"].append(
        ("DEMO-O1", "SYNTHETIC PERSON", "ΦΘ", "11:00", "Τρ-Πε", "DEMO-T"),
    ))
    summary = inspect_private_outpatient_source(source)
    assert summary.recurring_treatments == 3
    assert summary.weekday_placements == 7


def test_unknown_undocumented_header_or_missing_sheet_is_rejected(tmp_path):
    source = _fixture(tmp_path)
    _edit(source, lambda wb: setattr(
        wb["OUTPATIENT_SCHEDULE"]["F1"], "value", "Unknown",
    ))
    with pytest.raises(PrivateOutpatientSourceError, match="schema"):
        inspect_private_outpatient_source(source)
    source = _fixture(tmp_path)
    _edit(source, lambda wb: wb.remove(wb["OUTPATIENT_SCHEDULE"]))
    with pytest.raises(PrivateOutpatientSourceError, match="missing"):
        inspect_private_outpatient_source(source)


def test_empty_outpatient_schedule_is_not_a_full_real_source(tmp_path):
    source = _fixture(tmp_path, appointments=[])
    with pytest.raises(PrivateOutpatientSourceError, match="no appointments"):
        inspect_private_outpatient_source(source)


def test_unregistered_outpatient_patient_with_no_schedule_cannot_disappear(tmp_path):
    source = _fixture(tmp_path)
    _edit(source, lambda wb: wb["PATIENTS"].append(
        ("DEMO-O2", None, "SECOND DEMO", "Ο", "Ενεργός", "Εξωτερικός"),
    ))
    with pytest.raises(PrivateOutpatientSourceError, match="without approved appointments"):
        inspect_private_outpatient_source(source)


def test_invalid_source_does_not_expose_path_in_cli(tmp_path, capsys):
    source = tmp_path / "PRIVATE_REAL_PATIENT_HOSPITAL.xlsx"
    assert main(["--source", str(source)]) == 2
    text = capsys.readouterr().out
    assert "PRIVATE_REAL_PATIENT" not in text
    assert str(source) not in text


def test_audit_rejects_private_xlsx_inside_git_tree(tmp_path):
    source = _fixture(tmp_path)
    (tmp_path / ".git").mkdir()
    with pytest.raises(PrivateOutpatientSourceError, match="cannot reside in Git"):
        inspect_private_outpatient_source(source)


def test_actual_workbook_unchanged_under_failed_audit(tmp_path):
    source = _fixture(tmp_path)
    _edit(source, lambda wb: setattr(
        wb["OUTPATIENT_SCHEDULE"]["A2"], "value", "UNKNOWN",
    ))
    before = source.read_bytes()
    with pytest.raises(PrivateOutpatientSourceError):
        inspect_private_outpatient_source(source)
    assert source.read_bytes() == before
