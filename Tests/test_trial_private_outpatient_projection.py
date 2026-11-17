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


def _inpatient_projection(*, day="ΔΕ", clock=None, provider="DEMO T", approved_shared=False):
    from datetime import time
    from rehab_excel.trial_demo_source_projection import TrialDemoSourceProjection
    from rehab_excel.trial_patient_registry_projection import (
        PATIENTS_HEADERS, TrialPatientRegistryProjection,
    )
    from rehab_excel.trial_weekly_projection import TrialWeeklyProjection
    from rehab_excel.weekly_slots_source import HEADERS

    start = clock or time(9, 15)
    patients = [("IN-1", None, "FICTIONAL INPATIENT A", "Ο", "Ενεργός", "Εσωτερικός", None, None, None)]
    weekly = [("IN-1", "ΦΘ", day, start, provider, "C2")]
    if approved_shared:
        patients.append(("IN-2", None, "FICTIONAL INPATIENT B", "Ο", "Ενεργός", "Εσωτερικός", None, None, None))
        weekly.append(("IN-2", "ΦΘ", day, start, provider, "C3"))
    return TrialDemoSourceProjection(
        TrialPatientRegistryProjection(PATIENTS_HEADERS, tuple(patients), len(patients)),
        TrialWeeklyProjection(HEADERS, tuple(weekly), len(weekly)),
        len(patients), len(weekly),
    )


@pytest.mark.parametrize("provider", ["DEMO T", " demo t "])
def test_mixed_patient_types_cannot_share_provider_slot_without_approval(tmp_path, provider):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    source = _source(tmp_path)
    before = source.read_bytes()
    outpatient = project_private_outpatient_source(source)
    with pytest.raises(ValueError, match="explicit shared-slot approval") as err:
        merge_trial_demo_outpatients(_inpatient_projection(provider=provider), outpatient)
    assert "FICTIONAL" not in str(err.value)
    assert source.read_bytes() == before


@pytest.mark.parametrize("field", ["day", "clock", "provider"])
def test_mixed_sources_with_distinct_provider_slots_remain_complete(tmp_path, field):
    from datetime import time
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    options = {"day": "ΤΡ", "clock": time(9, 45), "provider": "OTHER DEMO T"}
    inpatient = _inpatient_projection(**{field: options[field]})
    outpatient = project_private_outpatient_source(_source(tmp_path))
    combined = merge_trial_demo_outpatients(inpatient, outpatient)
    assert combined.patients == 2
    assert combined.weekday_slot_count == 6
    assert combined.weekly_slots == inpatient.weekly_slots
    assert combined.outpatient is outpatient


def test_outpatient_pair_cannot_implicitly_share_one_provider_slot(tmp_path):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    source = _source(tmp_path)
    def add_peer(wb):
        wb["PATIENTS"].append(("OUT-2", None, "FICTIONAL PEER", "Ο", "Ενεργός", "Εξωτερικός"))
        wb["OUTPATIENT_SCHEDULE"].append(("OUT-2", "FICTIONAL PEER", "ΦΘ", "09:15", "Τε", "DEMO T"))
    _modify(source, add_peer)
    outpatient = project_private_outpatient_source(source)
    with pytest.raises(ValueError, match="explicit shared-slot approval"):
        merge_trial_demo_outpatients(_inpatient_projection(day="ΤΡ"), outpatient)


def test_existing_approved_inpatient_group_is_not_rejected_or_expanded(tmp_path):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    outpatient = project_private_outpatient_source(_source(tmp_path))
    existing = _inpatient_projection(day="ΤΡ", approved_shared=True)
    combined = merge_trial_demo_outpatients(existing, outpatient)
    assert combined.patients == 3
    assert combined.weekday_slot_count == 7
    assert combined.weekly_slots.rows == existing.weekly_slots.rows
    with pytest.raises(ValueError, match="explicit shared-slot approval"):
        merge_trial_demo_outpatients(
            _inpatient_projection(approved_shared=True), outpatient,
        )
