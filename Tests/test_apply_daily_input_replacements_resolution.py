from __future__ import annotations

import inspect
from datetime import date, time
from types import SimpleNamespace

import pytest

from rehab_core.models import Session, SessionCancellationKind
from tools.apply_daily_input_replacements import (
    SUSPEND_CHOICE,
    _append_treatment_suspension,
    _ask_resolution,
    _replacement_id,
    main,
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


def test_replacement_id_uses_legacy_first_generation_then_versions_history():
    session = Session(
        session_id="planner:57:ΦΘ@2026-10-19",
        patient_id="56",
        therapist_id="Σάρρας",
        session_date=date(2026, 10, 19),
        start_time=time(8, 30),
        treatment="ΦΘ",
    )
    base = "daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19"

    assert _replacement_id(session, ()) == base

    history_v1 = (SimpleNamespace(replacement_id=base),)
    assert _replacement_id(session, history_v1) == f"{base}:v2"

    history_v2 = (
        SimpleNamespace(replacement_id=base),
        SimpleNamespace(replacement_id=f"{base}:v2"),
        SimpleNamespace(replacement_id="daily-input:unrelated:2026-10-19"),
    )
    assert _replacement_id(session, history_v2) == f"{base}:v3"


def test_rerun_persists_only_new_replacement_facts_and_cleans_partial_output():
    source = inspect.getsource(main)

    assert "new_replacements: list[ReplacementAssignment] = []" in source
    assert "new_replacements.append(replacement)" in source
    assert "append_daily_replacement_facts(\n            output,\n            new_replacements," in source
    assert "if output_created and output.exists():" in source
    assert "output.unlink()" in source
