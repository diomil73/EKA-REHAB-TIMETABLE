import pytest
from openpyxl import Workbook

from rehab_core.models import PatientType
from rehab_excel.patient_registry_source import (
    PatientRegistrySourceError,
    read_patient_registry,
)


def _save(tmp_path, headers, rows):
    path = tmp_path / "patients.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "PATIENTS"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


def test_legacy_five_column_registry_defaults_to_inpatient(tmp_path):
    path = _save(
        tmp_path,
        ["PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση"],
        [["P1", "101", "ΑΣΘΕΝΗΣ", "Ο", "Ενεργός"]],
    )

    [patient] = read_patient_registry(path)

    assert patient.patient_type == PatientType.INPATIENT
    assert patient.hospital_mrn is None
    assert patient.room == "101"


def test_extended_registry_reads_outpatient_and_hospital_mrn(tmp_path):
    path = _save(
        tmp_path,
        [
            "PatientID",
            "Θάλαμος",
            "Ασθενής",
            "Λοιμώδης",
            "Κατάσταση",
            "ΤύποςΑσθενή",
            "HospitalMRN",
        ],
        [["O1", None, "ΕΞΩΤΕΡΙΚΟΣ", "Ο", None, "Εξωτερικός", "MRN-77"]],
    )

    [patient] = read_patient_registry(path)

    assert patient.patient_type == PatientType.OUTPATIENT
    assert patient.is_outpatient is True
    assert patient.hospital_mrn == "MRN-77"
    assert patient.room is None


@pytest.mark.parametrize(
    "label",
    ["Εξωτερικός", "ΕΞΩΤΕΡΙΚΟΣ", "Εξωτερικος"],
)
def test_greek_outpatient_labels_normalize_reliably(tmp_path, label):
    path = _save(
        tmp_path,
        [
            "PatientID",
            "Θάλαμος",
            "Ασθενής",
            "Λοιμώδης",
            "Κατάσταση",
            "ΤύποςΑσθενή",
        ],
        [["O1", None, "ΕΞΩΤΕΡΙΚΟΣ", "Ο", None, label]],
    )

    [patient] = read_patient_registry(path)
    assert patient.patient_type == PatientType.OUTPATIENT


def test_unknown_patient_type_is_rejected_instead_of_defaulting(tmp_path):
    path = _save(
        tmp_path,
        [
            "PatientID",
            "Θάλαμος",
            "Ασθενής",
            "Λοιμώδης",
            "Κατάσταση",
            "PatientType",
        ],
        [["X1", None, "UNKNOWN", "Ο", None, "visitor"]],
    )

    with pytest.raises(PatientRegistrySourceError, match="unknown patient type"):
        read_patient_registry(path)


def test_outpatient_infectious_combination_is_rejected_by_domain(tmp_path):
    path = _save(
        tmp_path,
        [
            "PatientID",
            "Θάλαμος",
            "Ασθενής",
            "Λοιμώδης",
            "Κατάσταση",
            "PatientType",
        ],
        [["O1", None, "ΕΞΩΤΕΡΙΚΟΣ", "Ν", None, "outpatient"]],
    )

    with pytest.raises(ValueError, match="cannot be marked infectious"):
        read_patient_registry(path)
