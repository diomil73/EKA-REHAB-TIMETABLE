from datetime import date, time
from types import SimpleNamespace

import pytest

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


@pytest.fixture(autouse=True)
def generalized_workbook(monkeypatch):
    settings = SimpleNamespace(
        therapies=("ΦΘ", "Ρομποτικό", "Ανακλινόμενο", "Πισίνα", "Εργο", "Λογο", "ΕΦΑ", "Ψυχ"),
        day_patterns=("Δε-Τε-Πα", "Τρ-Πε", "Καθ/να"),
        standard_timeslots=(time(9, 0), time(10, 0), time(11, 0)),
        reclined_timeslots=(time(8, 30), time(9, 30), time(10, 30), time(11, 30)),
    )
    provider_map = {
        "Εργο": ("OT1", "OT2"),
        "Λογο": ("SLP1",),
        "Ψυχ": ("PSY1",),
    }

    class Draft:
        def spec_for(self, treatment):
            providers = provider_map.get(treatment, ())
            return SimpleNamespace(
                treatment=treatment,
                provider_choices=providers,
                requires_provider=bool(providers),
            )

    monkeypatch.setattr(automatic, "read_settings", lambda *a, **k: settings)
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: ())
    monkeypatch.setattr(
        automatic,
        "build_patient_program_draft",
        lambda *a, **k: Draft(),
    )
    return settings


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


def test_generic_provider_treatment_is_scheduled(monkeypatch):
    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Εργο", "Δε-Τε-Πα"),),
    )

    assert proposal.complete is True
    assert proposal.assignments[0].treatment == "Εργο"
    assert proposal.assignments[0].therapist_name in {"OT1", "OT2"}
    assert proposal.assignments[0].start_time in {time(9, 0), time(10, 0), time(11, 0)}


def test_providerless_treatment_is_scheduled_with_blank_provider():
    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("ΕΦΑ", "Τρ-Πε"),),
    )

    assert proposal.complete is True
    assignment = proposal.assignments[0]
    assert assignment.therapist_id == ""
    assert assignment.therapist_name == ""


def test_missing_day_pattern_is_filled_from_settings(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(9, 0)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("ΦΘ"),),
    )

    assert proposal.complete is True
    assert proposal.assignments[0].day_pattern in {"Δε-Τε-Πα", "Τρ-Πε", "Καθ/να"}


def test_manual_provider_constraint_is_never_replaced(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (
            _choice("T1", "Άλφα", 1, time(9, 0)),
            _choice("T2", "Βήτα", 5, time(10, 0)),
        ),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("ΦΘ", "Δε-Τε-Πα", provider_id="Βήτα"),),
    )

    assert proposal.complete is True
    assert proposal.assignments[0].therapist_id == "T2"
    assert proposal.assignments[0].therapist_name == "Βήτα"


def test_manual_generic_provider_and_time_are_never_replaced():
    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest(
                "Εργο",
                "Δε-Τε-Πα",
                provider_id="OT2",
                start_time=time(11, 0),
            ),
        ),
    )

    assert proposal.complete is True
    assignment = proposal.assignments[0]
    assert assignment.therapist_name == "OT2"
    assert assignment.start_time == time(11, 0)


def test_unavailable_manual_constraint_is_explicitly_unresolved(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(9, 0)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("ΦΘ", "Δε-Τε-Πα", start_time=time(12, 0)),),
    )

    assert proposal.complete is False
    assert len(proposal.assignments) == 0
    assert "χειροκίνητες επιλογές" in proposal.unresolved[0].reason


def test_generic_provider_capacity_six_rejects_overloaded_day(monkeypatch):
    entries = tuple(
        SimpleNamespace(
            patient_id=f"P{i}",
            treatment="Εργο",
            start_time=time(8 + i, 0),
            day_pattern="Δε-Τε-Πα",
            therapist_id="OT1",
        )
        for i in range(6)
    )
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: entries)

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="NEW",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Εργο", "Δε-Τε-Πα", provider_id="OT1"),),
    )

    assert proposal.complete is False


def test_recliner_prefers_nearest_slot_to_physiotherapy(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(10, 0)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Δε-Τε-Πα"),
            AutomaticTreatmentRequest("Ανακλινόμενο", "Δε-Τε-Πα"),
        ),
    )

    assert proposal.complete is True
    by_treatment = {item.treatment: item for item in proposal.assignments}
    assert by_treatment["ΦΘ"].start_time == time(10, 0)
    assert by_treatment["Ανακλινόμενο"].start_time in {time(9, 30), time(10, 30)}


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
