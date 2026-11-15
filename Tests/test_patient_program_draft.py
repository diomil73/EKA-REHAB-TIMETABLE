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


def test_program_draft_requires_provider_only_for_physiotherapy_but_supports_optional_names(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    result = draft.build_patient_program_draft("dummy.xlsm", patient_id="P001")

    physio = result.spec_for("ΦΘ")
    assert physio.provider_kind == draft.PROVIDER_PHYSIOTHERAPIST
    assert physio.provider_choices == ("Φυσιο 1", "Φυσιο 2")
    assert physio.supports_provider is True
    assert physio.requires_provider is True

    robotic = result.spec_for("Ρομποτικό")
    assert robotic.provider_kind == draft.PROVIDER_PHYSIOTHERAPIST
    assert robotic.supports_provider is True
    assert robotic.requires_provider is False

    ergo = result.spec_for("Εργο")
    assert ergo.provider_kind == draft.PROVIDER_OCCUPATIONAL_THERAPIST
    assert ergo.provider_choices == ("Εργο 1",)
    assert ergo.supports_provider is True
    assert ergo.requires_provider is False

    logo = result.spec_for("Λογο")
    assert logo.provider_kind == draft.PROVIDER_SPEECH_THERAPIST
    assert logo.provider_choices == ("Λογο 1",)
    assert logo.supports_provider is True
    assert logo.requires_provider is False

    psych = result.spec_for("Ψυχ")
    assert psych.provider_kind == draft.PROVIDER_PSYCHOLOGIST
    assert psych.provider_choices == ("Ψυχ 1",)
    assert psych.supports_provider is True
    assert psych.requires_provider is False

    for treatment in ("Ανακλινόμενο", "Πισίνα", "ΕΦΑ"):
        spec = result.spec_for(treatment)
        assert spec.supports_provider is False
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


def test_program_draft_accepts_free_text_optional_provider_name(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    result = draft.build_patient_program_draft(
        "dummy.xlsm",
        patient_id="P001",
        constraints=(
            draft.TreatmentDraftConstraint(
                treatment="Εργο",
                provider_id="Χειροκίνητο Όνομα",
            ),
        ),
    )

    assert result.constraints[0].provider_id == "Χειροκίνητο Όνομα"


def test_program_draft_rejects_provider_on_treatment_without_name_field(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    with pytest.raises(ValueError, match="does not have a provider/name field"):
        draft.build_patient_program_draft(
            "dummy.xlsm",
            patient_id="P001",
            constraints=(
                draft.TreatmentDraftConstraint(
                    treatment="Πισίνα",
                    provider_id="Κάποιος",
                ),
            ),
        )


def test_program_draft_is_read_only_and_requires_known_patient(monkeypatch):
    monkeypatch.setattr(draft, "read_settings", lambda path: _settings())
    monkeypatch.setattr(draft, "read_patients", lambda path: [_patient()])

    with pytest.raises(ValueError, match="Unknown PatientID"):
        draft.build_patient_program_draft("dummy.xlsm", patient_id="P999")
