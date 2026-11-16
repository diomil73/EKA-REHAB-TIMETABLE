"""Fictional-only source projection regression. No real clinical records."""
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.trial_private_outpatient_projection import (
    TrialPrivateOutpatientProjectionError, project_private_outpatient_source,
)


def _source(tmp_path: Path) -> Path:
    workbook = Workbook()
    patients = workbook.active
    patients.title = "PATIENTS"
    patients.append((
        "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης",
        "Κατάσταση", "PatientType", "HospitalMRN",
        "ResponsibleDoctor", "SourceContext",
    ))
    patients.append((
        "OUT-1", None, "FICTIONAL PATIENT", "Ο",
        "Ενεργός", "Εξωτερικός", None, None, "APPROVED SYNTHETIC",
    ))
    schedule = workbook.create_sheet("OUTPATIENT_SCHEDULE")
    schedule.append((
        "PatientID", "Ασθενής", "Θεραπεία",
        "Ώρα", "Ημέρες", "Θεραπευτής",
    ))
    schedule.append(("OUT-1", "FICTIONAL PATIENT", "ΦΘ", "09:15", "Δε-Τε-Πα", "DEMO T"))
    schedule.append(("OUT-1", "FICTIONAL PATIENT", "Εργο", "10:15", "Τρ-Πε", None))
    path = tmp_path / "private_synthetic_outpatients.xlsx"
    workbook.save(path)
    workbook.close()
    return path


def _modify(path, updater):
    workbook = load_workbook(path)
    updater(workbook)
    workbook.save(path)
    workbook.close()


def test_projection_preserves_existing_application_fields_and_exact_weekdays(tmp_path):
    source = _source(tmp_path)
    original = source.read_bytes()
    result = project_private_outpatient_source(source)
    assert result.patient_count == 1
    assert result.appointment_count == 2
    assert result.weekday_slot_count == 5
    assert result.patient_rows[0][-1] == "APPROVED SYNTHETIC"
    assert result.appointment_rows[0][:3] == ("OUT-1", "FICTIONAL PATIENT", "ΦΘ")
    assert {r[3] for r in result.weekday_view_rows} == {"ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ"}
    assert all(r[6].startswith("OUTPATIENT_SCHEDULE:") for r in result.weekday_view_rows)
    assert source.read_bytes() == original
    assert "FICTIONAL PATIENT" not in repr(result)


def test_populated_unknown_metadata_fails_closed(tmp_path):
    source = _source(tmp_path)
    _modify(source, lambda wb: wb["PATIENTS"].cell(2, 10, "UNAPPROVED DETAIL"))
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)
    source = _source(tmp_path)
    _modify(source, lambda wb: wb["OUTPATIENT_SCHEDULE"].cell(2, 7, "PRIVATE NOTE"))
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)


def test_readonly_source_rejected_in_git_checkout(tmp_path):
    source = _source(tmp_path)
    (tmp_path / ".git").mkdir()
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)


def test_patient_type_mismatch_is_not_imported(tmp_path):
    source = _source(tmp_path)
    _modify(source, lambda wb: wb["PATIENTS"].cell(2, 6, "Εσωτερικός"))
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)
