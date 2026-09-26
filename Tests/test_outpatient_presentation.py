from datetime import date, time

import pytest

from rehab_core import (
    BaseScheduleEntry,
    DailySessionState,
    DailySessionStatus,
    Patient,
    PatientType,
    ReplacementProviderKind,
)
from rehab_excel.outpatient_presentation import (
    OutpatientPresentationError,
    outpatient_daily_targets,
    validate_outpatient_base_entries,
)


DAY = date(2026, 9, 21)


def state(
    session_id: str,
    patient_id: str,
    *,
    status: DailySessionStatus = DailySessionStatus.ACTIVE,
    therapist_id: str = "T1",
    at: time = time(9, 0),
    effective_therapist_id: str | None = None,
    effective_time: time | None = None,
    provider_kind: ReplacementProviderKind | None = None,
) -> DailySessionState:
    return DailySessionState(
        session_id=session_id,
        patient_id=patient_id,
        session_date=DAY,
        status=status,
        original_therapist_id=therapist_id,
        original_time=at,
        effective_therapist_id=(
            therapist_id
            if status == DailySessionStatus.ACTIVE
            else effective_therapist_id
        ),
        effective_time=at if status == DailySessionStatus.ACTIVE else effective_time,
        effective_provider_kind=(
            ReplacementProviderKind.THERAPIST
            if status == DailySessionStatus.ACTIVE
            else provider_kind
        ),
    )


def outpatient(patient_id: str = "P-OUT") -> Patient:
    return Patient(
        patient_id=patient_id,
        display_name="Εξωτερικός",
        patient_type=PatientType.OUTPATIENT,
    )


def inpatient(patient_id: str = "P-IN") -> Patient:
    return Patient(patient_id=patient_id, display_name="Εσωτερικός")


def test_active_outpatient_marks_only_its_effective_daily_slot_blue():
    target = outpatient_daily_targets(
        [state("S1", "P-OUT", therapist_id="T1", at=time(9, 0))],
        [outpatient()],
    )

    assert len(target) == 1
    assert target[0].provider_id == "T1"
    assert target[0].start_time == time(9, 0)
    assert target[0].patient_ids == ("P-OUT",)


def test_inpatient_does_not_create_outpatient_blue_target():
    targets = outpatient_daily_targets(
        [state("S1", "P-IN")],
        [inpatient()],
    )

    assert targets == ()


def test_replaced_outpatient_marks_replacement_destination_blue():
    targets = outpatient_daily_targets(
        [
            state(
                "S1",
                "P-OUT",
                status=DailySessionStatus.REPLACED,
                therapist_id="T1",
                at=time(9, 0),
                effective_therapist_id="T2",
                effective_time=time(10, 0),
                provider_kind=ReplacementProviderKind.THERAPIST,
            )
        ],
        [outpatient()],
    )

    assert len(targets) == 1
    assert targets[0].provider_id == "T2"
    assert targets[0].start_time == time(10, 0)


def test_cancelled_outpatient_releases_slot_and_does_not_claim_blue_active_target():
    targets = outpatient_daily_targets(
        [
            state(
                "S1",
                "P-OUT",
                status=DailySessionStatus.CANCELLED,
                therapist_id="T1",
                at=time(9, 0),
            )
        ],
        [outpatient()],
    )

    assert targets == ()


def test_same_recurring_slot_can_be_blue_on_outpatient_day_without_other_day_patient_affecting_it():
    # Daily states contain only the patient actually active on the concrete day.
    targets = outpatient_daily_targets(
        [state("S-OUT", "P-OUT", therapist_id="T1", at=time(11, 0))],
        [outpatient(), inpatient()],
    )

    assert len(targets) == 1
    assert targets[0].patient_ids == ("P-OUT",)


def test_mixed_inpatient_and_outpatient_in_same_effective_daily_cell_is_rejected():
    states = [
        state("S-OUT", "P-OUT", therapist_id="T1", at=time(11, 0)),
        state("S-IN", "P-IN", therapist_id="T1", at=time(11, 0)),
    ]

    with pytest.raises(OutpatientPresentationError, match="mixed inpatient/outpatient"):
        outpatient_daily_targets(states, [outpatient(), inpatient()])


def test_outpatient_robotic_flag_is_rejected():
    entry = BaseScheduleEntry(
        base_entry_id="B1",
        patient_id="P-OUT",
        treatment="ΦΘ",
        start_time=time(9, 0),
        day_pattern="Δε-Τε-Πα",
        therapist_id="T1",
        robotic=True,
    )

    with pytest.raises(OutpatientPresentationError, match="cannot have robotic"):
        validate_outpatient_base_entries([entry], [outpatient()])


def test_outpatient_robotic_treatment_label_is_rejected_even_if_flag_is_false():
    entry = BaseScheduleEntry(
        base_entry_id="B1",
        patient_id="P-OUT",
        treatment="Ρομποτικό",
        start_time=time(9, 0),
        day_pattern="Δε-Τε-Πα",
        therapist_id="T1",
        robotic=False,
    )

    with pytest.raises(OutpatientPresentationError, match="cannot have robotic"):
        validate_outpatient_base_entries([entry], [outpatient()])


def test_inpatient_robotic_assignment_remains_allowed():
    entry = BaseScheduleEntry(
        base_entry_id="B1",
        patient_id="P-IN",
        treatment="Ρομποτικό",
        start_time=time(9, 0),
        day_pattern="Δε-Τε-Πα",
        therapist_id="T1",
        robotic=True,
    )

    validate_outpatient_base_entries([entry], [inpatient()])
