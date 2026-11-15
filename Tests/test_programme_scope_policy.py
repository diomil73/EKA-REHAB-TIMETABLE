"""Synthetic-only regression checks for the independent afternoon/medical policy."""
from datetime import date, time

import pytest

from rehab_core.programme_scope_policy import (
    Absence, AbsenceReason, AbsenceScope, PlannedSession, Programme,
    SchedulingScopeError, affected_session_ids, conflicting_session_pairs,
)

DAY = date(2026, 11, 16)


def slot(sid, programme, start, end, *, patient="SYNTH-1", provider="SYNTH-PT", treatment="ΦΘ"):
    return PlannedSession(sid, patient, DAY, programme, treatment, time(*start), time(*end), provider)


def absence(reason, scope, **kwargs):
    return Absence("SYNTH-1", DAY, reason, scope, **kwargs)


def fixture_sessions():
    return (
        slot("AM", Programme.MORNING, (9, 0), (9, 30)),
        slot("POOL", Programme.MORNING, (10, 0), (10, 30), treatment="Πισίνα"),
        slot("PM", Programme.AFTERNOON, (16, 0), (16, 30), provider="SYNTH-PM-PT"),
        slot("OTHER", Programme.AFTERNOON, (16, 0), (16, 30), patient="SYNTH-2", provider="SYNTH-PT2"),
    )


@pytest.mark.parametrize("reason", [
    AbsenceReason.MEDICAL, AbsenceReason.TEMPORARY_DISCHARGE, AbsenceReason.TRANSFER,
])
def test_full_day_reason_covers_all_patient_treatments(reason):
    assert affected_session_ids(fixture_sessions(), absence(reason, AbsenceScope.FULL_DAY)) == (
        "AM", "POOL", "PM"
    )


def test_medical_window_impact_is_bounded_to_overlapping_sessions():
    a = absence(AbsenceReason.MEDICAL, AbsenceScope.TIME_WINDOW, start=time(9, 15), end=time(9, 45))
    assert affected_session_ids(fixture_sessions(), a) == ("AM",)


def test_medical_window_can_explicitly_target_afternoon():
    a = absence(
        AbsenceReason.MEDICAL, AbsenceScope.TIME_WINDOW,
        programme=Programme.AFTERNOON, start=time(16), end=time(17),
    )
    assert affected_session_ids(fixture_sessions(), a) == ("PM",)


def test_medical_window_uses_half_open_intervals():
    a = absence(AbsenceReason.MEDICAL, AbsenceScope.TIME_WINDOW, start=time(9, 30), end=time(10))
    assert affected_session_ids(fixture_sessions(), a) == ()


def test_specific_morning_physio_deferral_does_not_touch_afternoon():
    a = absence(
        AbsenceReason.SESSION_DEFERRAL, AbsenceScope.SESSION,
        programme=Programme.MORNING, session_id="AM",
    )
    assert affected_session_ids(fixture_sessions(), a) == ("AM",)


def test_afternoon_provider_absence_only_targets_afternoon_session():
    a = absence(
        AbsenceReason.THERAPIST_ABSENCE, AbsenceScope.SESSION,
        programme=Programme.AFTERNOON, session_id="PM",
    )
    assert affected_session_ids(fixture_sessions(), a) == ("PM",)


def test_session_target_must_match_patient_date_and_programme():
    with pytest.raises(SchedulingScopeError):
        affected_session_ids(fixture_sessions(), absence(
            AbsenceReason.MEDICAL, AbsenceScope.SESSION, session_id="OTHER"
        ))
    with pytest.raises(SchedulingScopeError):
        affected_session_ids(fixture_sessions(), absence(
            AbsenceReason.MEDICAL, AbsenceScope.SESSION,
            programme=Programme.AFTERNOON, session_id="AM",
        ))


def test_conflict_checks_ignore_cross_programme_overlap():
    a = slot("M", Programme.MORNING, (15, 0), (16, 0))
    b = slot("A", Programme.AFTERNOON, (15, 0), (16, 0))
    assert conflicting_session_pairs((a, b)) == ()


def test_conflict_checks_detect_same_programme_patient_and_provider():
    a = slot("M1", Programme.MORNING, (9, 0), (9, 30))
    b = slot("M2", Programme.MORNING, (9, 15), (9, 45), provider="SYNTH-OTHER")
    c = slot("M3", Programme.MORNING, (9, 15), (9, 45), patient="SYNTH-2")
    assert conflicting_session_pairs((a, b, c)) == (("M1", "M2"), ("M1", "M3"))


@pytest.mark.parametrize("treatment,provider", [
    ("Εργο", "SYNTH-OT"), ("Πισίνα", "SYNTH-PT"), ("ΦΘ", None),
])
def test_afternoon_is_physio_only_with_named_provider(treatment, provider):
    with pytest.raises(SchedulingScopeError):
        slot("PM", Programme.AFTERNOON, (16, 0), (16, 30), provider=provider, treatment=treatment)


@pytest.mark.parametrize("reason", [AbsenceReason.TRANSFER, AbsenceReason.TEMPORARY_DISCHARGE])
def test_global_patient_state_cannot_be_narrowed(reason):
    with pytest.raises(SchedulingScopeError):
        absence(reason, AbsenceScope.SESSION, session_id="AM")


def test_medical_requires_explicit_scope_and_valid_interval():
    with pytest.raises(SchedulingScopeError):
        absence(AbsenceReason.MEDICAL, AbsenceScope.TIME_WINDOW, start=time(10), end=time(10))
    with pytest.raises(SchedulingScopeError):
        absence(AbsenceReason.MEDICAL, AbsenceScope.FULL_DAY, programme=Programme.MORNING)


def test_patient_schedule_is_immutable_under_preview():
    rows = fixture_sessions()
    previous = tuple(rows)
    affected_session_ids(rows, absence(AbsenceReason.MEDICAL, AbsenceScope.FULL_DAY))
    assert rows == previous


def test_duplicate_ids_stop_before_conflict_or_cancellation():
    rows = (slot("D", Programme.MORNING, (9, 0), (10, 0)),
            slot("D", Programme.AFTERNOON, (16, 0), (17, 0)))
    with pytest.raises(SchedulingScopeError):
        conflicting_session_pairs(rows)
    with pytest.raises(SchedulingScopeError):
        affected_session_ids(rows, absence(AbsenceReason.MEDICAL, AbsenceScope.FULL_DAY))
