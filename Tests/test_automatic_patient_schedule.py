from datetime import date, time

from rehab_excel import automatic_patient_schedule as automatic
from rehab_excel.automatic_patient_schedule import (
    AutomaticTreatmentRequest,
    propose_automatic_patient_schedule,
)
from rehab_excel.patient_recurring_scheduling import RecurringTherapistChoice


DAY = date(2026, 10, 12)


def _choice(therapist_id, name, workload, *slots):
    return RecurringTherapistChoice(
        therapist_id=therapist_id,
        display_name=name,
        workload_count=workload,
        available_slots=tuple(slots),
    )


def test_solver_prefers_lower_workload_complete_proposal(monkeypatch):
    def fake_choices(*args, treatment, **kwargs):
        if treatment == "ΦΘ":
            return (
                _choice("T1", "Άλφα", 5, time(9, 0)),
                _choice("T2", "Βήτα", 2, time(10, 0)),
            )
        return (_choice("T3", "Γάμμα", 1, time(11, 0)),)

    monkeypatch.setattr(automatic, "recurring_therapist_choices", fake_choices)

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Δε-Τε-Πα"),
            AutomaticTreatmentRequest("Ρομποτικό", "Τρ-Πε"),
        ),
    )

    assert proposal.complete is True
    assert proposal.total_workload == 3
    assert [(item.treatment, item.therapist_id, item.start_time) for item in proposal.assignments] == [
        ("ΦΘ", "T2", time(10, 0)),
        ("Ρομποτικό", "T3", time(11, 0)),
    ]


def test_solver_avoids_same_patient_time_collision_on_overlapping_days(monkeypatch):
    def fake_choices(*args, treatment, **kwargs):
        if treatment == "ΦΘ":
            return (_choice("T1", "Άλφα", 1, time(9, 0)),)
        return (_choice("T2", "Βήτα", 1, time(9, 0), time(10, 0)),)

    monkeypatch.setattr(automatic, "recurring_therapist_choices", fake_choices)

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Δε-Τε-Πα"),
            AutomaticTreatmentRequest("Ρομποτικό", "Δε-Τε-Πα"),
        ),
    )

    assert proposal.complete is True
    assert {item.start_time for item in proposal.assignments} == {time(9, 0), time(10, 0)}


def test_same_time_is_allowed_when_day_patterns_do_not_overlap(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, treatment, **kwargs: (
            _choice("T1" if treatment == "ΦΘ" else "T2", treatment, 1, time(9, 0)),
        ),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Τρ-Πε"),
            AutomaticTreatmentRequest("Ρομποτικό", "Δε-Τε-Πα"),
        ),
    )

    assert proposal.complete is True
    assert [item.start_time for item in proposal.assignments] == [time(9, 0), time(9, 0)]


def test_unsupported_treatment_is_explicitly_unresolved(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(9, 0)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Καθ/να"),
            AutomaticTreatmentRequest("Εργο", "Καθ/να"),
        ),
    )

    assert proposal.complete is False
    assert len(proposal.assignments) == 1
    assert proposal.unresolved[0].treatment == "Εργο"
    assert "automatic physiotherapist solver" in proposal.unresolved[0].reason


def test_partial_proposal_when_all_supported_options_conflict(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, treatment, **kwargs: (
            _choice("T1" if treatment == "ΦΘ" else "T2", treatment, 1, time(9, 0)),
        ),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Καθ/να"),
            AutomaticTreatmentRequest("Ρομποτικό", "Καθ/να"),
        ),
    )

    assert len(proposal.assignments) == 1
    assert len(proposal.unresolved) == 1
    assert "συγκρούονται χρονικά" in proposal.unresolved[0].reason
