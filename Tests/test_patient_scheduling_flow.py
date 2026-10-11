from datetime import date, time

from rehab_core.scheduling_candidates import TherapistScheduleCandidate
from rehab_excel.patient_scheduling_flow import (
    TreatmentAssignment,
    available_slots_for_therapist,
    refresh_and_validate_assignment,
    therapist_choices,
    validate_assignment_against_plan,
)
from rehab_excel.patient_treatment_plan import (
    PatientTreatmentPlan,
    TreatmentSchedulingOptions,
)


DAY = date(2026, 10, 12)


def candidate(name: str, load: int, *slots: time) -> TherapistScheduleCandidate:
    return TherapistScheduleCandidate(
        therapist_id=name,
        display_name=name,
        active_sessions=load,
        active_timeslots=load,
        max_daily_timeslots=6,
        available_slots=tuple(slots),
    )


def plan() -> PatientTreatmentPlan:
    return PatientTreatmentPlan(
        patient_id="P001",
        target_date=DAY,
        treatments=(
            TreatmentSchedulingOptions(
                treatment="ΦΘ",
                requires_therapist=True,
                therapist_candidates=(
                    candidate("Γαύρας", 3, time(9, 0), time(10, 0)),
                    candidate("Άλλος", 4, time(11, 0)),
                ),
            ),
            TreatmentSchedulingOptions(
                treatment="Πισίνα",
                requires_therapist=False,
            ),
        ),
    )


def test_therapist_choices_preserve_engine_workload_order_and_label():
    choices = therapist_choices(plan(), "ΦΘ")

    assert [item.therapist_id for item in choices] == ["Γαύρας", "Άλλος"]
    assert choices[0].label == "Γαύρας 3"
    assert choices[1].label == "Άλλος 4"


def test_selected_therapist_exposes_only_engine_available_slots():
    slots = available_slots_for_therapist(plan(), "ΦΘ", "Γαύρας")

    assert slots == (time(9, 0), time(10, 0))
    assert available_slots_for_therapist(plan(), "ΦΘ", "missing") == ()


def test_valid_assignment_must_match_both_therapist_and_slot():
    current = plan()

    valid = validate_assignment_against_plan(
        current,
        TreatmentAssignment("ΦΘ", "Γαύρας", time(10, 0)),
    )
    invalid_time = validate_assignment_against_plan(
        current,
        TreatmentAssignment("ΦΘ", "Γαύρας", time(11, 0)),
    )
    invalid_therapist = validate_assignment_against_plan(
        current,
        TreatmentAssignment("ΦΘ", "Nobody", time(10, 0)),
    )

    assert valid.valid is True
    assert invalid_time.valid is False
    assert "time" in invalid_time.message.lower()
    assert invalid_therapist.valid is False
    assert "therapist" in invalid_therapist.message.lower()


def test_non_therapist_treatment_is_not_silently_routed_through_this_flow():
    result = validate_assignment_against_plan(
        plan(),
        TreatmentAssignment("Πισίνα", "Γαύρας", time(10, 0)),
    )

    assert result.valid is False
    assert "not therapist-scheduled" in result.message


def test_refresh_rebuilds_authoritative_plan_before_commit(monkeypatch):
    refreshed = PatientTreatmentPlan(
        patient_id="P001",
        target_date=DAY,
        treatments=(
            TreatmentSchedulingOptions(
                treatment="ΦΘ",
                requires_therapist=True,
                therapist_candidates=(candidate("Γαύρας", 4, time(9, 0)),),
            ),
        ),
    )

    monkeypatch.setattr(
        "rehab_excel.patient_scheduling_flow.build_patient_treatment_plan",
        lambda *args, **kwargs: refreshed,
    )

    result = refresh_and_validate_assignment(
        "dummy.xlsm",
        patient_id="P001",
        selected_treatments=["ΦΘ"],
        target_date=DAY,
        assignment=TreatmentAssignment("ΦΘ", "Γαύρας", time(10, 0)),
    )

    assert result.valid is False
    assert result.refreshed_plan is refreshed
    assert "no longer available" in result.message
