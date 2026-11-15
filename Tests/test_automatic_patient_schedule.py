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

    class Draft:
        def spec_for(self, treatment):
            return SimpleNamespace(
                treatment=treatment,
                provider_choices=(),
                requires_provider=False,
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
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (
            _choice("T1", "Άλφα", 5, time(9, 0)),
            _choice("T2", "Βήτα", 2, time(10, 0)),
        ),
    )

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
    by_treatment = {item.treatment: item for item in proposal.assignments}
    assert by_treatment["ΦΘ"].therapist_id == "T2"
    assert by_treatment["ΦΘ"].start_time == time(10, 0)
    assert by_treatment["Ρομποτικό"].therapist_id == ""
    assert by_treatment["Ρομποτικό"].start_time == time(9, 0)


def test_solver_avoids_same_patient_interval_collision_on_overlapping_days(monkeypatch):
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
        lambda *args, **kwargs: (_choice("T1", "ΦΘ", 1, time(9, 0)),),
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


def test_ergo_is_scheduled_with_blank_provider():
    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Εργο", "Δε-Τε-Πα"),),
    )

    assert proposal.complete is True
    assignment = proposal.assignments[0]
    assert assignment.treatment == "Εργο"
    assert assignment.therapist_id == ""
    assert assignment.therapist_name == ""


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


def test_manual_time_for_non_physio_treatment_is_preserved():
    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest(
                "Εργο",
                "Δε-Τε-Πα",
                start_time=time(11, 0),
            ),
        ),
    )

    assert proposal.complete is True
    assignment = proposal.assignments[0]
    assert assignment.therapist_name == ""
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


def test_robotic_does_not_require_provider(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(10, 0)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Ρομποτικό", "Δε-Τε-Πα"),),
    )

    assert proposal.complete is True
    assignment = proposal.assignments[0]
    assert assignment.therapist_id == ""
    assert assignment.therapist_name == ""


def test_recliner_prefers_duration_safe_adjacent_slot_to_physiotherapy(monkeypatch):
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
    # 09:30-10:00 is exactly adjacent. 10:30 overlaps the 10:00-10:45 ΦΘ.
    assert by_treatment["Ανακλινόμενο"].start_time == time(9, 30)


def test_physio_0915_forces_recliner_to_1000_or_0830_and_prefers_1000(monkeypatch, generalized_workbook):
    generalized_workbook.reclined_timeslots = (
        time(8, 30),
        time(9, 0),
        time(9, 30),
        time(10, 0),
    )
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(9, 15)),),
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
    assert by_treatment["ΦΘ"].start_time == time(9, 15)
    assert by_treatment["Ανακλινόμενο"].start_time == time(10, 0)


def test_recliner_0900_conflicts_with_physio_0915(monkeypatch, generalized_workbook):
    generalized_workbook.reclined_timeslots = (time(9, 0),)
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(9, 15)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Δε-Τε-Πα", start_time=time(9, 15)),
            AutomaticTreatmentRequest("Ανακλινόμενο", "Δε-Τε-Πα", start_time=time(9, 0)),
        ),
    )

    assert len(proposal.assignments) == 1
    assert len(proposal.unresolved) == 1
    reason = proposal.unresolved[0].reason
    assert "συγκρού" in reason
    assert "09:00-09:30" in reason
    assert "09:15-10:00" in reason
    assert "Δευτέρα" in reason


def test_partial_proposal_when_fixed_options_conflict(monkeypatch):
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "ΦΘ", 1, time(9, 0)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Καθ/να", start_time=time(9, 0)),
            AutomaticTreatmentRequest("Ρομποτικό", "Καθ/να", start_time=time(9, 0)),
        ),
    )

    assert len(proposal.assignments) == 1
    assert len(proposal.unresolved) == 1
    assert "συγκρού" in proposal.unresolved[0].reason


def test_edit_draft_ignores_old_physio_when_new_recliner_is_adjacent(monkeypatch, generalized_workbook):
    generalized_workbook.standard_timeslots = (time(11, 30),)
    generalized_workbook.reclined_timeslots = (time(11, 0),)

    old_physio = SimpleNamespace(
        patient_id="P1",
        treatment="ΦΘ",
        start_time=time(11, 0),
        day_pattern="Δε-Τε-Πα",
        therapist_id="Άλφα",
    )
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: (old_physio,))
    monkeypatch.setattr(
        automatic,
        "recurring_therapist_choices",
        lambda *args, **kwargs: (_choice("T1", "Άλφα", 1, time(11, 30)),),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest("ΦΘ", "Καθ/να", start_time=time(11, 30)),
            AutomaticTreatmentRequest("Ανακλινόμενο", "Δε-Τε-Πα", start_time=time(11, 0)),
        ),
    )

    assert proposal.complete is True
    by_treatment = {item.treatment: item for item in proposal.assignments}
    assert by_treatment["Ανακλινόμενο"].start_time == time(11, 0)
    assert by_treatment["ΦΘ"].start_time == time(11, 30)


def test_manual_physio_busy_slot_names_day_provider_and_patient(monkeypatch):
    occupied = SimpleNamespace(
        patient_id="P2",
        treatment="ΦΘ",
        start_time=time(11, 30),
        day_pattern="Δε-Τε-Πα",
        therapist_id="Σαρράς",
    )
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: (occupied,))
    monkeypatch.setattr(automatic, "recurring_therapist_choices", lambda *a, **k: ())
    monkeypatch.setattr(
        automatic,
        "read_patients",
        lambda *a, **k: (
            SimpleNamespace(patient_id="P1", display_name="Τρέχων Ασθενής"),
            SimpleNamespace(patient_id="P2", display_name="Άλλος Ασθενής"),
        ),
    )

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(
            AutomaticTreatmentRequest(
                "ΦΘ",
                "Δε-Τε-Πα",
                provider_id="Σαρράς",
                start_time=time(11, 30),
            ),
        ),
    )

    assert proposal.complete is False
    reason = proposal.unresolved[0].reason
    assert "Δευτέρα 11:30" in reason
    assert "Σαρράς" in reason
    assert "Άλλος Ασθενής" in reason
    assert "ΦΘ" in reason
