from __future__ import annotations

from datetime import date, time

import pytest

from rehab_core.models import Session, SessionCancellationKind
from tools.apply_daily_input_replacements import (
    SUSPEND_CHOICE,
    _append_treatment_suspension,
    _ask_resolution,
)


def test_resolution_accepts_candidate_rank(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: "2")

    assert _ask_resolution(3) == 2


def test_resolution_accepts_explicit_treatment_suspension(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: "s")

    assert _ask_resolution(3) == SUSPEND_CHOICE


def test_resolution_with_no_candidates_still_requires_suspend_or_abort(monkeypatch):
    answers = iter(["0", "x", "S"])
    monkeypatch.setattr("builtins.input", lambda _: next(answers))

    assert _ask_resolution(0) == SUSPEND_CHOICE


def test_resolution_abort_raises_keyboard_interrupt(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: "Q")

    with pytest.raises(KeyboardInterrupt):
        _ask_resolution(2)


def test_append_treatment_suspension_creates_final_date_specific_fact():
    session = Session(
        session_id="S1",
        patient_id="P1",
        therapist_id="T1",
        session_date=date(2026, 10, 19),
        start_time=time(9, 15),
        treatment="ΦΘ",
    )
    cancellations = []
    suspensions = []

    suspension = _append_treatment_suspension(
        session,
        cancellations=cancellations,
        suspensions=suspensions,
    )

    assert cancellations == [suspension]
    assert suspensions == [suspension]
    assert suspension.target_session_id == "S1"
    assert suspension.cancellation_date == date(2026, 10, 19)
    assert suspension.kind == SessionCancellationKind.TREATMENT_SUSPENDED
    assert suspension.reason == "Αναστολή θεραπείας"
