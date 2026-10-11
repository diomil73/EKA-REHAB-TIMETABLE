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
        therapies=("Εργο", "Λογο"),
        day_patterns=("Δε-Τε-Πα", "Τρ-Πε", "Καθ/να"),
        standard_timeslots=(time(9, 0), time(10, 0)),
        reclined_timeslots=(),
    )


def test_generic_provider_failure_reports_missing_settings_pool(monkeypatch):
    class Draft:
        def spec_for(self, treatment):
            return SimpleNamespace(
                treatment=treatment,
                requires_provider=True,
                provider_choices=(),
            )

    monkeypatch.setattr(automatic, "read_settings", lambda *a, **k: _settings())
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: ())
    monkeypatch.setattr(automatic, "build_patient_program_draft", lambda *a, **k: Draft())

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="P1",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Εργο"),),
    )

    assert proposal.complete is False
    assert "SETTINGS" in proposal.unresolved[0].reason
    assert "providers" in proposal.unresolved[0].reason


def test_generic_provider_failure_reports_capacity_or_occupancy(monkeypatch):
    class Draft:
        def spec_for(self, treatment):
            return SimpleNamespace(
                treatment=treatment,
                requires_provider=True,
                provider_choices=("OT1",),
            )

    entries = tuple(
        SimpleNamespace(
            patient_id=f"P{i}",
            treatment="Εργο",
            start_time=time(8 + i, 0),
            day_pattern="Καθ/να",
            therapist_id="OT1",
        )
        for i in range(6)
    )

    monkeypatch.setattr(automatic, "read_settings", lambda *a, **k: _settings())
    monkeypatch.setattr(automatic, "read_base_schedule", lambda *a, **k: entries)
    monkeypatch.setattr(automatic, "build_patient_program_draft", lambda *a, **k: Draft())

    proposal = propose_automatic_patient_schedule(
        "dummy.xlsm",
        patient_id="NEW",
        anchor_date=DAY,
        requests=(AutomaticTreatmentRequest("Εργο", "Καθ/να"),),
    )

    assert proposal.complete is False
    assert "κανένας δεν έχει διαθέσιμο slot" in proposal.unresolved[0].reason
