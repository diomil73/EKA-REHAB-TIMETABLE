"""Fail-closed wiring tests for Excel's real global replacement pipeline.

These tests are synthetic and do not open or save an .xlsm.
"""
from datetime import date, time
from pathlib import Path

import pytest

from rehab_core.models import BaseScheduleEntry, Session
from rehab_core.patient_schedule import validate_daily_replacement_patient_slot


MONDAY = date(2026, 11, 16)
THURSDAY = date(2026, 11, 19)


def _entry(entry_id, day, clock, treatment, provider=None):
    return BaseScheduleEntry(
        base_entry_id=entry_id, patient_id="DEMO-P",
        treatment=treatment, start_time=clock, day_pattern=day,
        therapist_id=provider,
    )


def _target(day):
    return Session(
        session_id=f"weekly:2@{day.isoformat()}",
        patient_id="DEMO-P", therapist_id="DEMO-T1",
        session_date=day, start_time=time(9, 15),
        treatment="ΦΘ",
    )


def test_final_excel_gate_blocks_other_specialty_thursday_only():
    entries = [
        _entry("weekly:2", "Δε-Πε", time(9, 15), "ΦΘ", "DEMO-T1"),
        _entry("weekly:3", "Πε", time(12, 15), "Εργο"),
    ]
    validate_daily_replacement_patient_slot(
        entries=entries, target_session=_target(MONDAY),
        replacement_time=time(12, 15),
    )
    with pytest.raises(ValueError, match="Patient slot blocked"):
        validate_daily_replacement_patient_slot(
            entries=entries, target_session=_target(THURSDAY),
            replacement_time=time(12, 15),
        )


def test_same_treatment_different_weekday_time_remains_independent():
    entries = [
        _entry("weekly:2", "Δε", time(9, 15), "ΦΘ", "DEMO-T1"),
        _entry("weekly:3", "Πε", time(12, 15), "ΦΘ", "DEMO-T2"),
    ]
    validate_daily_replacement_patient_slot(
        entries=entries, target_session=_target(MONDAY),
        replacement_time=time(12, 15),
    )
    with pytest.raises(ValueError, match="Patient slot blocked"):
        validate_daily_replacement_patient_slot(
            entries=entries, target_session=_target(THURSDAY),
            replacement_time=time(12, 15),
        )


def test_operational_global_path_forwards_effective_weekly_source():
    """Guard against a future refactor silently bypassing the safety argument."""
    root = Path(__file__).resolve().parents[1]
    global_code = (root / "Python/rehab_core/global_replacement_end_to_end.py").read_text(encoding="utf-8")
    policy_code = (root / "Python/rehab_core/replacement_policy.py").read_text(encoding="utf-8")
    cli_code = (root / "Python/tools/prepare_global_replacements_in_place.py").read_text(encoding="utf-8")
    assert "base_entries=base_entries_tuple" in global_code
    assert "base_entries=base_entries" in policy_code
    assert "validate_daily_replacement_patient_slot(" in cli_code
    assert "replacement_time=assignment.replacement_time" in cli_code
