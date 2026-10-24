from datetime import date, time

import pytest
from openpyxl import Workbook

from rehab_core.base_schedule import materialize_sessions_for_date
from rehab_core.models import Patient, PatientType
from rehab_core.workload import calculate_therapist_workload
from rehab_excel.outpatient_schedule_source import (
    OutpatientScheduleSourceError,
    read_outpatient_schedule,
    read_unified_base_schedule,
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


def test_outpatient_recurring_entry_counts_once_in_therapist_workload(tmp_path):
    path = _workbook(
        tmp_path,
        rows=[("O1", "ΕΞΩΤΕΡΙΚΟΣ", "ΦΘ", "09:15", "Δε-Τε-Πα", "T1")],
    )
    patient = outpatient()
    target_date = date(2026, 10, 19)  # Monday

    entries = read_outpatient_schedule(path, patients=[patient])
    sessions = materialize_sessions_for_date(entries, target_date)
    workload = calculate_therapist_workload(
        "T1",
        target_date,
        sessions,
        patients=[patient],
    )

    assert len(sessions) == 1
    assert workload.active_sessions == 1
    assert workload.active_timeslots == 1


def test_unified_schedule_reuses_supplied_patient_registry(monkeypatch):
    patients = [outpatient()]
    monkeypatch.setattr(
        "rehab_excel.outpatient_schedule_source.read_patient_registry",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("unexpected patient reread")),
    )
    monkeypatch.setattr(
        "rehab_excel.outpatient_schedule_source.read_base_schedule",
        lambda *_args, **_kwargs: ["inpatient-entry"],
    )
    monkeypatch.setattr(
        "rehab_excel.outpatient_schedule_source.read_outpatient_schedule",
        lambda *_args, patients=None, **_kwargs: [patients[0].patient_id],
    )

    assert read_unified_base_schedule("book.xlsm", patients=patients) == [
        "inpatient-entry",
        "O1",
    ]


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
