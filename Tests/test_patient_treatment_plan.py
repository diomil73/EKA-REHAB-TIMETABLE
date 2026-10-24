from datetime import date, time
from types import SimpleNamespace

import pytest

from rehab_core.scheduling_candidates import TherapistScheduleCandidate
from rehab_excel.patient_treatment_plan import build_patient_treatment_plan


DAY = date(2026, 10, 12)


def _settings(*therapies: str):
    return SimpleNamespace(therapies=therapies)


def test_requires_at_least_one_selected_treatment(monkeypatch):
    monkeypatch.setattr(
        "rehab_excel.patient_treatment_plan.read_settings",
        lambda path: _settings("ΦΘ"),
    )

    with pytest.raises(ValueError, match="At least one treatment"):
        build_patient_treatment_plan(
            "dummy.xlsm",
            patient_id="P001",
            selected_treatments=[],
            target_date=DAY,
        )


def test_rejects_duplicate_treatment_selection_case_insensitively(monkeypatch):
    monkeypatch.setattr(
        "rehab_excel.patient_treatment_plan.read_settings",
        lambda path: _settings("ΦΘ"),
    )

    with pytest.raises(ValueError, match="Duplicate treatment"):
        build_patient_treatment_plan(
            "dummy.xlsm",
            patient_id="P001",
            selected_treatments=["ΦΘ", " φθ "],
            target_date=DAY,
        )


def test_rejects_treatment_not_defined_in_settings(monkeypatch):
    monkeypatch.setattr(
        "rehab_excel.patient_treatment_plan.read_settings",
        lambda path: _settings("ΦΘ", "Ρομποτικό"),
    )

    with pytest.raises(ValueError, match="not defined in SETTINGS"):
        build_patient_treatment_plan(
            "dummy.xlsm",
            patient_id="P001",
            selected_treatments=["Πισίνα"],
            target_date=DAY,
        )


def test_physio_treatment_reuses_ranked_candidate_engine(monkeypatch):
    candidate = TherapistScheduleCandidate(
        therapist_id="T1",
        display_name="Γαύρας",
        active_sessions=3,
        active_timeslots=3,
        max_daily_timeslots=6,
        available_slots=(time(9, 0), time(10, 0)),
    )
    calls = []

    monkeypatch.setattr(
        "rehab_excel.patient_treatment_plan.read_settings",
        lambda path: _settings("ΦΘ", "Πισίνα"),
    )

    def fake_candidates(path, **kwargs):
        calls.append((path, kwargs))
        return (candidate,)

    monkeypatch.setattr(
        "rehab_excel.patient_treatment_plan.workbook_therapist_candidates",
        fake_candidates,
    )

    plan = build_patient_treatment_plan(
        "dummy.xlsm",
        patient_id="P001",
        selected_treatments=["ΦΘ", "Πισίνα"],
        target_date=DAY,
    )

    assert plan.selected_treatments == ("ΦΘ", "Πισίνα")
    assert plan.treatments[0].requires_therapist is True
    assert plan.treatments[0].therapist_candidates == (candidate,)
    assert plan.treatments[1].requires_therapist is False
    assert plan.treatments[1].therapist_candidates == ()
    assert calls == [
        (
            "dummy.xlsm",
            {
                "patient_id": "P001",
                "treatment": "ΦΘ",
                "target_date": DAY,
                "additional_robotic_capable_names": (),
            },
        )
    ]


def test_settings_spelling_is_preserved_as_canonical_label(monkeypatch):
    monkeypatch.setattr(
        "rehab_excel.patient_treatment_plan.read_settings",
        lambda path: _settings("Ρομποτικό"),
    )
    monkeypatch.setattr(
        "rehab_excel.patient_treatment_plan.workbook_therapist_candidates",
        lambda *args, **kwargs: (),
    )

    plan = build_patient_treatment_plan(
        "dummy.xlsm",
        patient_id="P001",
        selected_treatments=[" ρομποτικό "],
        target_date=DAY,
    )

    assert plan.selected_treatments == ("Ρομποτικό",)
    assert plan.treatments[0].requires_therapist is True
