from pathlib import Path

import pytest

from rehab_core.models import Patient, PatientType
from Python.tools.audit_master_infectious_coverage import (
    InfectiousCoverageError, count_master_infectious,
)


def test_aggregate_count_excludes_outpatients_and_names(tmp_path):
    source = tmp_path / "demo.xlsm"
    source.write_bytes(b"synthetic")
    patients = [
        Patient("A", "Synthetic 1", infectious=True),
        Patient("B", "Synthetic 2", infectious=False),
        Patient("C", "Synthetic 3", patient_type=PatientType.OUTPATIENT),
    ]
    result = count_master_infectious(source, reader=lambda _: patients)
    assert result == {
        "total_inpatients": 2,
        "infectious_inpatients": 1,
        "noninfectious_inpatients": 1,
    }
    assert "Synthetic" not in str(result)


def test_refuse_nonexistent_or_non_xlsm(tmp_path):
    with pytest.raises(InfectiousCoverageError, match="Existing demo"):
        count_master_infectious(tmp_path / "not_here.xlsm")
    source = tmp_path / "demo.xlsx"
    source.write_bytes(b"synthetic")
    with pytest.raises(InfectiousCoverageError, match="Existing demo"):
        count_master_infectious(source)


def test_source_read_exception_redacted(tmp_path):
    source = tmp_path / "demo.xlsm"
    source.write_bytes(b"synthetic")
    def broken(_):
        raise RuntimeError("PRIVATE PATIENT RECORD")
    with pytest.raises(InfectiousCoverageError) as caught:
        count_master_infectious(source, reader=broken)
    assert "PRIVATE PATIENT" not in str(caught.value)


def test_empty_registry_refuses_acceptance(tmp_path):
    source = tmp_path / "demo.xlsm"
    source.write_bytes(b"synthetic")
    with pytest.raises(InfectiousCoverageError, match="No inpatient"):
        count_master_infectious(source, reader=lambda _: [])
