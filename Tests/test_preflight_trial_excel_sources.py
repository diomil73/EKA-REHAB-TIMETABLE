"""No patient records in two-source pilot preflight tests."""
from pathlib import Path
from types import SimpleNamespace

import pytest

from tools import preflight_trial_excel_sources as gate


def _sources(tmp_path):
    review = tmp_path / "candidate.xlsx"
    cc = tmp_path / "protected.xlsm"
    review.write_bytes(b"synthetic-review")
    cc.write_bytes(b"synthetic-master")
    return review, cc


def _review(ready=False):
    return SimpleNamespace(
        ready=ready, candidate_patients=3, verified_patient_ids=0 if not ready else 3,
        weekly_proposals=12, unadmitted_weekly_rows=0, unresolved_source_rows=1 if not ready else 0,
        overlap_rows=0, blockers=("synthetic hold",) if not ready else (),
    )


def _master(issues=()):
    return {
        "issues": tuple(issues), "operational_sheets_found": 4,
        "history_rows": 0,
    }


def test_preflight_stops_on_unverified_patient_ids(tmp_path, monkeypatch):
    review, cc = _sources(tmp_path)
    before = review.read_bytes(), cc.read_bytes()
    monkeypatch.setattr(gate, "evaluate_trial_admission", lambda path: _review())
    monkeypatch.setattr(gate, "inspect_source", lambda path: _master())
    status = gate.inspect_native_pilot_inputs(review=review, cc=cc)
    assert not status["native_trial_permitted"]
    assert status["verified_patient_ids"] == 0
    assert status["inputs_unchanged"]
    assert (review.read_bytes(), cc.read_bytes()) == before


def test_preflight_stops_on_historical_cc_issues_even_if_review_ready(tmp_path, monkeypatch):
    review, cc = _sources(tmp_path)
    monkeypatch.setattr(gate, "evaluate_trial_admission", lambda path: _review(True))
    monkeypatch.setattr(
        gate, "inspect_source",
        lambda path: _master(("EXISTING_TEST_OR_UNKNOWN_HISTORY_REQUIRES_ISOLATION",)),
    )
    status = gate.inspect_native_pilot_inputs(review=review, cc=cc)
    assert status["review_gate_ready"]
    assert not status["cc_intake_ready"]
    assert not status["native_trial_permitted"]


def test_preflight_pass_still_does_not_save_or_import(tmp_path, monkeypatch):
    review, cc = _sources(tmp_path)
    before = review.read_bytes(), cc.read_bytes()
    monkeypatch.setattr(gate, "evaluate_trial_admission", lambda path: _review(True))
    monkeypatch.setattr(gate, "inspect_source", lambda path: _master())
    status = gate.inspect_native_pilot_inputs(review=review, cc=cc)
    assert status["native_trial_permitted"]
    assert (review.read_bytes(), cc.read_bytes()) == before


def test_preflight_refuses_missing_or_same_workbook(tmp_path):
    review, cc = _sources(tmp_path)
    with pytest.raises(ValueError, match="distinct"):
        gate.inspect_native_pilot_inputs(review=review, cc=review)
    cc.unlink()
    with pytest.raises(ValueError, match="missing"):
        gate.inspect_native_pilot_inputs(review=review, cc=cc)


def test_preflight_aborts_if_external_source_changes_during_scan(tmp_path, monkeypatch):
    review, cc = _sources(tmp_path)
    monkeypatch.setattr(gate, "evaluate_trial_admission", lambda path: _review(True))

    def concurrent_update(path):
        cc.write_bytes(b"externally updated while scanning")
        return _master()

    monkeypatch.setattr(gate, "inspect_source", concurrent_update)
    with pytest.raises(RuntimeError, match="input changed"):
        gate.inspect_native_pilot_inputs(review=review, cc=cc)
