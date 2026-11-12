"""Synthetic regression tests for stable inpatient identity through relocations."""
from datetime import date, time

import pytest

from rehab_core.inpatient_placement import (
    InpatientPlacementChange, InpatientPlacementError,
    patient_with_current_placement, placement_on_date,
)
from rehab_core.models import BaseScheduleEntry, Patient, PatientType
from rehab_core.base_schedule import materialize_sessions_for_date


START = date(2026, 11, 1)
ROOM_DAY = date(2026, 11, 12)
DOCTOR_DAY = date(2026, 11, 16)


def inpatient():
    return Patient(
        patient_id="DEMO-P1", display_name="SYNTHETIC",
        room="A1", responsible_doctor="DR-A",
    )


def events():
    return (
        InpatientPlacementChange("DEMO-P1", ROOM_DAY, new_room="B3"),
        InpatientPlacementChange("DEMO-P1", DOCTOR_DAY,
                                new_responsible_doctor="DR-B"),
    )


def test_room_and_doctor_change_have_distinct_effective_dates():
    patient = inpatient()
    assert placement_on_date(
        patient, baseline_date=START, target_date=date(2026, 11, 11),
        changes=events(),
    ).room == "A1"
    mid = placement_on_date(
        patient, baseline_date=START, target_date=ROOM_DAY, changes=events()
    )
    assert (mid.room, mid.responsible_doctor) == ("B3", "DR-A")
    later = placement_on_date(
        patient, baseline_date=START, target_date=DOCTOR_DAY, changes=events()
    )
    assert (later.room, later.responsible_doctor) == ("B3", "DR-B")
    assert patient.room == "A1"
    assert patient.responsible_doctor == "DR-A"


def test_patient_id_and_future_therapies_survive_room_change():
    patient = inpatient()
    original_entry = BaseScheduleEntry(
        base_entry_id="weekly:2", patient_id=patient.patient_id,
        treatment="ΦΘ", start_time=time(9, 15),
        day_pattern="Καθ/να", therapist_id="T1",
    )
    changed = patient_with_current_placement(
        patient, baseline_date=START, target_date=DOCTOR_DAY, changes=events()
    )
    assert changed.patient_id == patient.patient_id
    assert changed.display_name == patient.display_name
    assert changed.room == "B3"
    assert changed.responsible_doctor == "DR-B"
    for day in [date(2026, 11, 11), date(2026, 11, 16)]:
        sessions = materialize_sessions_for_date([original_entry], day)
        assert len(sessions) == 1
        assert sessions[0].patient_id == "DEMO-P1"
        assert sessions[0].therapist_id == "T1"
    assert original_entry.patient_id == "DEMO-P1"


def test_change_of_only_doctor_does_not_reset_room():
    record = InpatientPlacementChange(
        "DEMO-P1", DOCTOR_DAY, new_responsible_doctor="DR-C"
    )
    state = placement_on_date(
        inpatient(), baseline_date=START,
        target_date=DOCTOR_DAY, changes=[record],
    )
    assert state.room == "A1"
    assert state.responsible_doctor == "DR-C"


@pytest.mark.parametrize("change", [
    InpatientPlacementChange("DIFFERENT", ROOM_DAY, new_room="B3"),
    InpatientPlacementChange("DEMO-P1", date(2026, 10, 30), new_room="B3"),
])
def test_wrong_patient_or_event_before_baseline_is_rejected(change):
    with pytest.raises(InpatientPlacementError):
        placement_on_date(
            inpatient(), baseline_date=START,
            target_date=DOCTOR_DAY, changes=[change],
        )


def test_same_day_conflicting_updates_fail_closed():
    with pytest.raises(InpatientPlacementError, match="same day"):
        placement_on_date(
            inpatient(), baseline_date=START, target_date=DOCTOR_DAY,
            changes=[
                InpatientPlacementChange("DEMO-P1", ROOM_DAY, new_room="B3"),
                InpatientPlacementChange("DEMO-P1", ROOM_DAY,
                                         new_responsible_doctor="DR-B"),
            ],
        )


def test_outpatient_cannot_be_given_inpatient_ward_transition():
    outpatient = Patient(
        patient_id="DEMO-O", display_name="SYNTHETIC",
        patient_type=PatientType.OUTPATIENT,
    )
    with pytest.raises(InpatientPlacementError, match="inpatient"):
        placement_on_date(
            outpatient, baseline_date=START,
            target_date=DOCTOR_DAY, changes=(),
        )


@pytest.mark.parametrize("kwargs", [
    {"new_room": ""},
    {"new_responsible_doctor": " "},
    {},
])
def test_empty_placement_change_rejected(kwargs):
    with pytest.raises(InpatientPlacementError):
        InpatientPlacementChange("DEMO-P1", ROOM_DAY, **kwargs)
