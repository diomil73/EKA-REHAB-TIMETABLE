from datetime import date, time
from types import SimpleNamespace

from rehab_excel import automatic_patient_schedule as automatic
from rehab_excel.automatic_patient_schedule import (
    AutomaticTreatmentRequest,
    propose_automatic_patient_schedule,
)


DAY = date(2026, 10, 12)


def _settings():
    return SimpleNamespace(
        therapies=("Ρομποτικό", "Εργο", "Λογο", "Ψυχ"),
        day_patterns=("Δε-Τε-Πα", "Τρ-Πε", "Καθ/να"),
        standard_timeslots=(time(9, 0), time(10, 0)),
        reclined_timeslots=(),
    )


class ProviderlessDraft:
    def spec_for(self, treatment):
        return SimpleNamespace(
            treatment=treatment,
            requires_provider=False,
            provider_choices=(),
        )


def test_specialty_treatment_does_not_fail_when_provider_pool_is_empty(monkeypatch):
    monkeypatch.setattr(automatic, "read_settings", lambda *a, **k: _settings())
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: ())
    monkeypatch.setattr(automatic, "build_patient_program_draft", lambda *a, **k: ProviderlessDraft())

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Εργο"),),
    )

    assert proposal.complete is True
    assert proposal.assignments[0].therapist_name == ""


def test_manual_time_conflict_reports_patient_collision(monkeypatch):
    existing = SimpleNamespace(
        patient_id="P1",
        treatment="Λογο",
        start_time=time(9, 0),
        day_pattern="Καθ/να",
        therapist_id="",
    )
    monkeypatch.setattr(automatic, "read_settings", lambda *a, **k: _settings())
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: (existing,))
    monkeypatch.setattr(automatic, "build_patient_program_draft", lambda *a, **k: ProviderlessDraft())

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Εργο", "Καθ/να", start_time=time(9, 0)),),
    )

    assert proposal.complete is False
    assert "συγκρούεται" in proposal.unresolved[0].reason
