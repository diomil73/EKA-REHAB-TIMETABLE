"""Synthetic-only patient-management draft regression tests."""
from datetime import time

import pytest

from rehab_core.patient_programme_draft import (
    AFTERNOON_TREATMENT,
    POOL_TREATMENT,
    AfternoonPhysioDraft,
    PatientProgrammeDraftError,
    TreatmentProgramme,
    prepare_afternoon_physio_drafts,
    validate_treatment_programme,
)


def test_empty_fields_do_not_create_afternoon_sessions():
    assert prepare_afternoon_physio_drafts() == ()
    assert prepare_afternoon_physio_drafts(clock="", days=" ", therapist=None) == ()
    assert AfternoonPhysioDraft().blank


def test_explicit_afternoon_physio_monday_and_tuesday():
    entries = prepare_afternoon_physio_drafts(
        clock="16:30", days="ΔΕ-ΤΡ", therapist="SYNTH-PT-PM",
    )
    assert entries == (
        AfternoonPhysioDraft(time(16, 30), "ΔΕ", "SYNTH-PT-PM"),
        AfternoonPhysioDraft(time(16, 30), "ΤΡ", "SYNTH-PT-PM"),
    )
    assert all(not entry.blank for entry in entries)


@pytest.mark.parametrize("kwargs", [
    {"clock": "16:00", "days": "ΔΕ"},
    {"clock": "16:00", "therapist": "SYNTH-PT"},
    {"days": "ΔΕ", "therapist": "SYNTH-PT"},
    {"clock": "16:00", "days": "ΔΕ", "therapist": ""},
    {"clock": "16:00", "days": "", "therapist": "SYNTH-PT"},
])
def test_partial_afternoon_triples_fail_closed(kwargs):
    with pytest.raises(PatientProgrammeDraftError):
        prepare_afternoon_physio_drafts(**kwargs)


@pytest.mark.parametrize("clock,days,provider", [
    ("not-time", "ΔΕ", "SYNTH-PT"),
    ("16:00", "INVALID", "SYNTH-PT"),
    ("16:00/17:00", "ΔΕ-ΤΡ", "SYNTH-PT"),
    ("16:00", "ΔΕ", "SYNTH-PT1/SYNTH-PT2"),
])
def test_malformed_afternoon_proposals_do_not_guess(clock, days, provider):
    with pytest.raises(PatientProgrammeDraftError):
        prepare_afternoon_physio_drafts(clock=clock, days=days, therapist=provider)


def test_morning_pool_is_permitted_but_afternoon_pool_rejected():
    assert validate_treatment_programme(POOL_TREATMENT, TreatmentProgramme.MORNING) is None
    with pytest.raises(PatientProgrammeDraftError):
        validate_treatment_programme(POOL_TREATMENT, TreatmentProgramme.AFTERNOON)


def test_afternoon_only_accepts_physiotherapy():
    assert validate_treatment_programme(AFTERNOON_TREATMENT, TreatmentProgramme.AFTERNOON) is None
    for unsupported in ("Εργο", "Λογο", "Ρομποτικό", "Ψυχ", "ΕΦΑ", "Ανακλινόμενο"):
        with pytest.raises(PatientProgrammeDraftError):
            validate_treatment_programme(unsupported, TreatmentProgramme.AFTERNOON)


def test_draft_cannot_be_partially_instantiated():
    with pytest.raises(PatientProgrammeDraftError):
        AfternoonPhysioDraft(time(16, 0), "ΔΕ", None)


def test_invalid_day_cannot_enter_patient_draft():
    with pytest.raises(PatientProgrammeDraftError):
        AfternoonPhysioDraft(time(16, 0), "ΚΥ", "SYNTH-PT")

def test_afternoon_accepts_multiple_explicit_weekday_combinations():
    entries = prepare_afternoon_physio_drafts(
        clock="16:30",
        days="ΔΕ-ΤΡ/ΠΕ-ΠΑ",
        therapist="SYNTH-PT-PM",
    )
    assert tuple(entry.weekday for entry in entries) == ("ΔΕ", "ΤΡ", "ΠΕ", "ΠΑ")
    assert all(entry.start_time == time(16, 30) for entry in entries)
    assert all(entry.physiotherapist == "SYNTH-PT-PM" for entry in entries)


def test_afternoon_rejects_repeated_weekday_across_combinations():
    with pytest.raises(PatientProgrammeDraftError):
        prepare_afternoon_physio_drafts(
            clock="16:30",
            days="ΔΕ-ΤΡ/ΤΡ-ΠΑ",
            therapist="SYNTH-PT-PM",
        )
