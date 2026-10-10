from datetime import time

import pytest
from openpyxl import Workbook

from rehab_core.models import Patient, PatientType
from rehab_excel.outpatient_schedule_source import (
    OutpatientScheduleSourceError,
    read_outpatient_schedule,
)


def _workbook(tmp_path, *, with_sheet=True, rows=()):
    path = tmp_path / "source.xlsx"
    wb = Workbook()
    wb.active.title = "PATIENTS"
    if with_sheet:
        ws = wb.create_sheet("OUTPATIENT_SCHEDULE")
        ws.append(["PatientID", "Ασθενής", "Θεραπεία", "Ώρα", "Ημέρες", "Θεραπευτής"])
        for row in rows:
            ws.append(list(row))
    wb.save(path)
    return path


def outpatient(pid="O1", name="ΕΞΩΤΕΡΙΚΟΣ"):
    return Patient(
        patient_id=pid,
        display_name=name,
        patient_type=PatientType.OUTPATIENT,
    )


def test_missing_outpatient_schedule_is_backward_compatible(tmp_path):
    path = _workbook(tmp_path, with_sheet=False)

    assert read_outpatient_schedule(path, patients=[outpatient()]) == []


def test_reads_one_row_per_outpatient_recurring_treatment(tmp_path):
    path = _workbook(
        tmp_path,
        rows=[("O1", "ΕΞΩΤΕΡΙΚΟΣ", "ΦΘ", "09:15", "Δε-Τε-Πα", "T1")],
    )

    entries = read_outpatient_schedule(path, patients=[outpatient()])

    assert len(entries) == 1
    entry = entries[0]
    assert entry.base_entry_id == "outpatient:2"
    assert entry.patient_id == "O1"
    assert entry.treatment == "ΦΘ"
    assert entry.start_time == time(9, 15)
    assert entry.day_pattern == "Δε-Τε-Πα"
    assert entry.therapist_id == "T1"
    assert entry.robotic is False


def test_outpatient_schedule_rejects_inpatient_identity(tmp_path):
    path = _workbook(
        tmp_path,
        rows=[("P1", "ΕΣΩΤΕΡΙΚΟΣ", "ΦΘ", "10:00", "Καθ/να", "T1")],
    )
    inpatient = Patient(patient_id="P1", display_name="ΕΣΩΤΕΡΙΚΟΣ")

    with pytest.raises(OutpatientScheduleSourceError, match="references inpatient"):
        read_outpatient_schedule(path, patients=[inpatient])


def test_outpatient_schedule_rejects_robotic_treatment_label(tmp_path):
    path = _workbook(
        tmp_path,
        rows=[("O1", "ΕΞΩΤΕΡΙΚΟΣ", "Ρομποτικό", "11:00", "Τρ-Πε", "T1")],
    )

    with pytest.raises(Exception, match="cannot have robotic treatment"):
        read_outpatient_schedule(path, patients=[outpatient()])
