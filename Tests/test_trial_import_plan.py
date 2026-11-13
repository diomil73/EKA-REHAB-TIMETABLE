"""Synthetic-only tests: no clinical facts or writable Excel operations."""
from datetime import time
from types import SimpleNamespace

import pytest

from rehab_core.programme_mapping import SourceWeeklyEntry
from rehab_excel import trial_import_plan as planner


def _sources(tmp_path):
    review = tmp_path / "review.xlsx"
    intake = tmp_path / "intake.xlsm"
    review.write_bytes(b"synthetic-reviewed-records")
    intake.write_bytes(b"synthetic-intake")
    return review, intake


def _entry():
    return SourceWeeklyEntry("DEMO-1", "ΦΘ", "ΔΕ", time(11, 30),
                             "DEMO THERAPIST", "C4")


def _enable(monkeypatch, *, blocked=False, mismatch=False):
    def inspect(**kwargs):
        return {
            "native_trial_permitted": not blocked,
            "review_patients": 1, "verified_patient_ids": 1,
            "proposed_weekly_slots": 2 if mismatch else 1,
        }

    monkeypatch.setattr(planner, "inspect_native_pilot_inputs", inspect)
    monkeypatch.setattr(
        planner, "evaluate_trial_admission",
        lambda path: SimpleNamespace(
            ready=True, admitted_entries=(_entry(),),
            candidate_patients=1, verified_patient_ids=1,
        ),
    )


def test_approved_plan_has_immutable_hashes_and_exact_slots(tmp_path, monkeypatch):
    review, intake = _sources(tmp_path)
    original = review.read_bytes(), intake.read_bytes()
    _enable(monkeypatch)
    plan = planner.prepare_trial_import_plan(review=review, intake=intake)
    assert plan.patient_count == 1
    assert plan.weekday_slot_count == 1
    assert plan.entries == (_entry(),)
    assert "DEMO-1" not in repr(plan)
    plan.verify_sources(review, intake)
    assert (review.read_bytes(), intake.read_bytes()) == original


def test_blocked_admission_never_returns_plan(tmp_path, monkeypatch):
    review, intake = _sources(tmp_path)
    _enable(monkeypatch, blocked=True)
    with pytest.raises(planner.TrialImportPlanError, match="blocked"):
        planner.prepare_trial_import_plan(review=review, intake=intake)


def test_cross_check_rejects_inconsistent_readbacks(tmp_path, monkeypatch):
    review, intake = _sources(tmp_path)
    _enable(monkeypatch, mismatch=True)
    with pytest.raises(planner.TrialImportPlanError, match="changed"):
        planner.prepare_trial_import_plan(review=review, intake=intake)


def test_hash_bound_plan_rejects_stale_review_or_intake(tmp_path, monkeypatch):
    review, intake = _sources(tmp_path)
    _enable(monkeypatch)
    plan = planner.prepare_trial_import_plan(review=review, intake=intake)
    review.write_bytes(b"review-tampered")
    with pytest.raises(planner.TrialImportPlanError, match="stale"):
        plan.verify_sources(review, intake)
    review.write_bytes(b"synthetic-reviewed-records")
    intake.write_bytes(b"intake-tampered")
    with pytest.raises(planner.TrialImportPlanError, match="stale"):
        plan.verify_sources(review, intake)


def test_plan_rejects_same_or_missing_inputs(tmp_path):
    review, intake = _sources(tmp_path)
    with pytest.raises(planner.TrialImportPlanError, match="Separate"):
        planner.prepare_trial_import_plan(review=review, intake=review)
    intake.unlink()
    with pytest.raises(planner.TrialImportPlanError, match="missing"):
        planner.prepare_trial_import_plan(review=review, intake=intake)


def test_concurrent_input_change_aborts(tmp_path, monkeypatch):
    review, intake = _sources(tmp_path)

    def mutate(**kwargs):
        intake.write_bytes(b"unexpected-mutated-input")
        return {"native_trial_permitted": False}

    monkeypatch.setattr(planner, "inspect_native_pilot_inputs", mutate)
    with pytest.raises(planner.TrialImportPlanError, match="changed during planning"):
        planner.prepare_trial_import_plan(review=review, intake=intake)
