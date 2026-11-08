from datetime import date
from types import SimpleNamespace

import pytest

from tools.audit_therapist_daily_capacity import capacity_anomalies


def item(*, capacity=6, slots=6, sessions=8, planned=6,
         planned_slots=6, name="Κοσμίδης", shared=2, replacements=2):
    return SimpleNamespace(
        service_date=date(2026, 10, 30),
        therapist_id=name,
        capacity_timeslots=capacity,
        final_occupied_timeslots=slots,
        final_active_sessions=sessions,
        planned_sessions=planned,
        planned_occupied_timeslots=planned_slots,
        approved_shared_or_over_capacity_events=shared,
        replacements_received=replacements,
    )


def test_eight_patients_do_not_automatically_mean_eight_occupied_slots():
    row = item(slots=6, sessions=8)
    result = capacity_anomalies([row], {(row.service_date, row.therapist_id.casefold()): 8})
    assert len(result) == 1
    assert result[0]["capacity_slots"] == 6
    assert result[0]["final_occupied_slots"] == 6
    assert result[0]["delivered_sessions"] == 8
    assert result[0]["unique_patients"] == 8
    assert "OCCUPIED_SLOTS_EXCEED_CAPACITY" not in result[0]["flags"]
    assert "MULTIPLE_DISTINCT_PATIENTS_PER_SLOT" in result[0]["flags"]
    assert "DELIVERED_EXCEEDS_INITIAL_PLAN" in result[0]["flags"]


def test_exceeding_real_slot_capacity_remains_explicit():
    row = item(slots=8, sessions=8, shared=0)
    result = capacity_anomalies([row], {(row.service_date, row.therapist_id.casefold()): 8})
    assert "OCCUPIED_SLOTS_EXCEED_CAPACITY" in result[0]["flags"]
    assert "MULTIPLE_SESSIONS_PER_SLOT" not in result[0]["flags"]


def test_unique_patient_key_matches_production_casefold_contract():
    row = item(slots=6, sessions=8)
    counts = {(row.service_date, row.therapist_id.casefold()): 8}
    assert counts[(row.service_date, "κοσμίδης".casefold())] == 8
    assert capacity_anomalies([row], counts)[0]["unique_patients"] == 8


def test_missing_capacity_is_unverified_not_six_by_assumption():
    row = item(capacity=None, slots=5, sessions=8)
    result = capacity_anomalies([row], {})
    assert "CAPACITY_NOT_CONFIGURED" in result[0]["flags"]


def test_duplicate_facts_fail_closed():
    row = item()
    with pytest.raises(ValueError, match="Duplicate"):
        capacity_anomalies([row, row], {})


def test_requested_person_is_filtered_and_all_days_included():
    other = item(name="Αργέντος", slots=4, sessions=4, planned=4)
    selected = item(slots=4, sessions=4, planned=4)
    reports = capacity_anomalies([other, selected], {}, therapist="κοσμίδης")
    assert len(reports) == 1
    assert reports[0]["therapist"] == "Κοσμίδης"
    assert reports[0]["flags"] == ()


def test_strict_capacity_gate_retains_real_occupied_slot_violation():
    row = item(slots=9, sessions=9, planned=6, shared=0)
    result = capacity_anomalies([row], {(row.service_date, row.therapist_id.casefold()): 9})
    assert len(result) == 1
    assert "OCCUPIED_SLOTS_EXCEED_CAPACITY" in result[0]["flags"]
    assert result[0]["final_occupied_slots"] == 9
    assert result[0]["capacity_slots"] == 6


def test_diagnostic_cli_exposes_opt_in_fail_closed_capacity_gate():
    from pathlib import Path
    source = (
        Path(__file__).resolve().parents[1]
        / "Python" / "tools" / "audit_therapist_daily_capacity.py"
    ).read_text(encoding="utf-8")
    assert '"--strict-capacity"' in source
    assert '"OCCUPIED_SLOTS_EXCEED_CAPACITY" in row["flags"]' in source
    assert '"CAPACITY_NOT_CONFIGURED" in row["flags"]' in source
    assert 'print(f"SAFETY STOP: {len(blocking)} therapist/day capacity exceptions")' in source
    assert 'return 3' in source
