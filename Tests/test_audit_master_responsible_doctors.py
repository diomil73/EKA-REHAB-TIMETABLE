from pathlib import Path

import pytest

from rehab_core.models import Patient, PatientType
from Python.tools.audit_master_responsible_doctors import (
    DoctorCoverageError, responsible_doctor_coverage,
)


def test_doctor_coverage_counts_inpatients_without_exposing_names(tmp_path):
    file = tmp_path / "demo.xlsm"
    file.write_bytes(b"synthetic")
    patients = [
        Patient("1", "Fictional One", responsible_doctor="Dr Example"),
        Patient("2", "Fictional Two", responsible_doctor="  "),
        Patient("3", "Fictional Three", responsible_doctor=None),
        Patient("4", "External", patient_type=PatientType.OUTPATIENT,
                responsible_doctor="Dr External"),
    ]
    result = responsible_doctor_coverage(file, reader=lambda _: patients)
    assert result == {
        "total_inpatients": 3,
        "inpatients_with_doctor": 1,
        "inpatients_without_doctor": 2,
    }
    assert "Dr Example" not in str(result)
    assert "Fictional One" not in str(result)


def test_all_missing_doctors_reported_without_guessing(tmp_path):
    file = tmp_path / "demo.xlsm"
    file.write_bytes(b"synthetic")
    patients = [Patient("1", "Fictional A"), Patient("2", "Fictional B")]
    assert responsible_doctor_coverage(file, reader=lambda _: patients) == {
        "total_inpatients": 2,
        "inpatients_with_doctor": 0,
        "inpatients_without_doctor": 2,
    }


def test_missing_source_stops_before_registry_access(tmp_path):
    with pytest.raises(DoctorCoverageError, match="Existing demo"):
        responsible_doctor_coverage(
            tmp_path / "missing.xlsm",
            reader=lambda _: (_ for _ in ()).throw(AssertionError("unexpected read")),
        )


def test_registry_error_is_redacted(tmp_path):
    file = tmp_path / "demo.xlsm"
    file.write_bytes(b"synthetic")
    def broken(_):
        raise ValueError("SECRET PATIENT IDENTIFIER")
    with pytest.raises(DoctorCoverageError) as caught:
        responsible_doctor_coverage(file, reader=broken)
    assert "SECRET" not in str(caught.value)


def test_empty_inpatient_registry_fails_closed(tmp_path):
    file = tmp_path / "demo.xlsm"
    file.write_bytes(b"synthetic")
    with pytest.raises(DoctorCoverageError, match="No inpatient"):
        responsible_doctor_coverage(file, reader=lambda _: [
            Patient("1", "External", patient_type=PatientType.OUTPATIENT),
        ])
