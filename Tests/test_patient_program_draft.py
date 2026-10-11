from datetime import time
from types import SimpleNamespace

import pytest

from rehab_excel import patient_program_draft as draft


def _settings():
    return SimpleNamespace(
        therapies=("ΦΘ", "Ρομποτικό", "Ανακλινόμενο", "Πισίνα", "Εργο", "Λογο", "ΕΦΑ", "Ψυχ"),
        day_patterns=("Καθ/να", "Δε-Τε-Πα", "Τρ-Πε"),
        therapist_names=("Φυσιο 1", "Φυσιο 2"),
        occupational_therapist_names=("Εργο 1",),
        speech_therapist_names=("Λογο 1",),
        psychologist_names=("Ψυχ 1",),
    )


def _patient():
    return SimpleNamespace(patient_id="P001")


def test_program_draft_exposes_every_settings_treatment_in_order(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    result = draft.build_patient_program_draft("dummy.xlsm", patient_id="P001")

    assert result.treatment_names == _settings().therapies


def test_program_draft_requires_provider_only_for_physiotherapy(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    result = draft.build_patient_program_draft("dummy.xlsm", patient_id="P001")

    assert result.spec_for("ΦΘ").provider_kind == draft.PROVIDER_PHYSIOTHERAPIST
    assert result.spec_for("ΦΘ").provider_choices == ("Φυσιο 1", "Φυσιο 2")
    for treatment in ("Ρομποτικό", "Ανακλινόμενο", "Πισίνα", "Εργο", "Λογο", "ΕΦΑ", "Ψυχ"):
        spec = result.spec_for(treatment)
        assert spec.requires_provider is False
        assert spec.provider_choices == ()


def test_program_draft_preserves_partial_manual_constraints(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    result = draft.build_patient_program_draft(
        "dummy.xlsm",
        patient_id="P001",
        constraints=(
            draft.TreatmentDraftConstraint(
                treatment="φθ",
                day_pattern="δε-τε-πα",
            ),
            draft.TreatmentDraftConstraint(
                treatment="Ρομποτικό",
                start_time=time(11, 30),
            ),
        ),
    )

    assert result.constraints[0] == draft.TreatmentDraftConstraint(
        treatment="ΦΘ",
        day_pattern="Δε-Τε-Πα",
    )
    assert result.constraints[1] == draft.TreatmentDraftConstraint(
        treatment="Ρομποτικό",
        start_time=time(11, 30),
    )


def test_program_draft_rejects_provider_outside_physiotherapy(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    with pytest.raises(ValueError, match="does not use a provider"):
        draft.build_patient_program_draft(
            "dummy.xlsm",
            patient_id="P001",
            constraints=(
                draft.TreatmentDraftConstraint(
                    treatment="Ρομποτικό",
                    provider_id="Φυσιο 1",
                ),
            ),
        )


def test_program_draft_is_read_only_and_requires_known_patient(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    with pytest.raises(ValueError, match="Unknown PatientID"):
        draft.build_patient_program_draft("dummy.xlsm", patient_id="P999")
