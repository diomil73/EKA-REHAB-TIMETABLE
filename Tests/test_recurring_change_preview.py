"""Synthetic-only tests for read-only Excel recurring change bridge."""
from datetime import date, time
from pathlib import Path
from types import SimpleNamespace

import pytest

from rehab_core.models import BaseScheduleEntry
from rehab_core.recurring_therapy_changes import RecurringTherapyChange
from rehab_excel import recurring_change_preview as bridge


BASELINE = date(2026, 11, 2)
MOVE_DAY = date(2026, 11, 16)
BASE = BaseScheduleEntry("weekly:2", "SYN-P", "ΦΘ", time(9, 15), "Δε-Πε", "T-4")
OTHER = BaseScheduleEntry("weekly:3", "OTHER-P", "ΦΘ", time(11, 15), "Πε", "T-6")
CAPACITIES = {"T-4": 4, "T-6": 6}


def fixture_workbook(tmp_path, monkeypatch, *, entries=(BASE, OTHER),
                     names=("T-4", "T-6")):
    path = tmp_path / "SYNTHETIC_DEMO.xlsm"
    path.write_bytes(b"PRIVATE SYNTHETIC SOURCE")
    monkeypatch.setattr(bridge, "read_effective_base_entries", lambda p: tuple(entries))
    monkeypatch.setattr(
        bridge, "read_settings",
        lambda p: SimpleNamespace(therapist_names=names),
    )
    return path


def proposed(**kwargs):
    return RecurringTherapyChange(
        "weekly:2", "SYN-P", MOVE_DAY, **kwargs,
    )


def test_readonly_bridge_accepts_safe_move_and_preserves_workbook(tmp_path, monkeypatch):
    path = fixture_workbook(tmp_path, monkeypatch)
    before = path.read_bytes()
    response = bridge.preview_recurring_change_from_excel(
        path, baseline_date=BASELINE,
        change=proposed(new_therapist_id="T-6", new_start_time=time(12, 15)),
        expected_entry=BASE, provider_max_daily_slots=CAPACITIES,
    )
    assert response.base_entry_count == 2
    assert response.workbook_sha256
    assert response.proposed.effective_entries[0].therapist_id == "T-6"
    assert path.read_bytes() == before


def test_bridge_rejects_stale_row_derived_assignment_identity(tmp_path, monkeypatch):
    # A row insert can change weekly:2 while retaining the same row number.
    path = fixture_workbook(
        tmp_path, monkeypatch,
        entries=(BaseScheduleEntry(
            "weekly:2", "REPLACED-P", "ΦΘ", time(9, 15), "Δε", "T-4",
        ),),
    )
    with pytest.raises(bridge.RecurringExcelPreviewError, match="row identity shifted"):
        bridge.preview_recurring_change_from_excel(
            path, baseline_date=BASELINE, change=proposed(new_start_time=time(12, 15)),
            expected_entry=BASE, provider_max_daily_slots=CAPACITIES,
        )


@pytest.mark.parametrize("capacities", [
    {"T-4": 6},  # unknown therapist capacity
    {"T-4": 4, "T-6": 0},  # impossible zero
    {"T-4": 4, "T-6": 6, "FAKE-T": 6},  # phantom registry
    {"T-4": True, "T-6": 6},  # bool must not become numeric capacity
])
def test_bridge_fails_closed_on_unverified_capacity(tmp_path, monkeypatch, capacities):
    path = fixture_workbook(tmp_path, monkeypatch)
    with pytest.raises(bridge.RecurringExcelPreviewError, match="capacity"):
        bridge.preview_recurring_change_from_excel(
            path, baseline_date=BASELINE, change=proposed(new_start_time=time(12, 15)),
            expected_entry=BASE, provider_max_daily_slots=capacities,
        )


def test_bridge_respects_four_slot_therapist_limit(tmp_path, monkeypatch):
    extras = tuple(
        BaseScheduleEntry(f"weekly:{n + 10}", f"P-{n}", "ΦΘ",
                          time(hour, 15), "Δε", "T-4")
        for n, hour in enumerate((7, 8, 10, 11), start=1)
    )
    # Retain BASE as the edited source and supply 4 other distinct slots.
    path = fixture_workbook(tmp_path, monkeypatch, entries=(BASE, *extras, OTHER))
    from rehab_core.recurring_therapy_changes import RecurringTherapyChangeError
    with pytest.raises(RecurringTherapyChangeError, match="capacity"):
        bridge.preview_recurring_change_from_excel(
            path, baseline_date=BASELINE,
            change=proposed(new_start_time=time(12, 15)),
            expected_entry=BASE, provider_max_daily_slots=CAPACITIES,
        )


def test_bridge_detects_workbook_changed_during_inspection(tmp_path, monkeypatch):
    path = fixture_workbook(tmp_path, monkeypatch)
    def external_change(p):
        path.write_bytes(b"EXTERNAL CONCURRENT MUTATION")
        return (BASE, OTHER)
    monkeypatch.setattr(bridge, "read_effective_base_entries", external_change)
    with pytest.raises(bridge.RecurringExcelPreviewError, match="changed during"):
        bridge.preview_recurring_change_from_excel(
            path, baseline_date=BASELINE, change=proposed(new_start_time=time(12, 15)),
            expected_entry=BASE, provider_max_daily_slots=CAPACITIES,
        )


def test_optional_weekly_source_replaces_planner_without_double_book(monkeypatch, tmp_path):
    path = tmp_path / "SYN_DEMO.xlsm"
    path.write_bytes(b"SYNTHETIC")
    registry = [SimpleNamespace(patient_id="SYN-P", is_outpatient=False)]
    monkeypatch.setattr(bridge, "read_patient_registry", lambda p: registry)
    legacy = BaseScheduleEntry("planner:2", "SYN-P", "ΦΘ", time(9, 15), "Δε", "T-4")
    monkeypatch.setattr(bridge, "read_unified_base_schedule",
                        lambda p, patients: [legacy])
    weekly = BaseScheduleEntry("weekly:2", "SYN-P", "ΦΘ", time(10, 15), "Δε", "T-4")
    monkeypatch.setattr(bridge, "read_weekly_slots_sheet",
                        lambda ws, inpatient_ids: (weekly,))

    class FakeBook:
        sheetnames = ["PROGRAM_WEEKLY_SLOTS"]
        def __getitem__(self, name):
            return object()
        def close(self):
            self.closed = True
    fake = FakeBook()
    monkeypatch.setattr(bridge, "load_workbook", lambda *a, **k: fake)
    result = bridge.read_effective_base_entries(path)
    assert result == (weekly,)
    assert fake.closed



def test_persistent_ready_preview_requires_stable_assignment_id(tmp_path, monkeypatch):
    path = fixture_workbook(tmp_path, monkeypatch)
    from rehab_excel.stable_recurring_ids import StableRecurringIdentityError
    monkeypatch.setattr(
        bridge, "bind_stable_assignment_ids",
        lambda path, entries: (_ for _ in ()).throw(
            StableRecurringIdentityError("No immutable IDs in source")
        ),
    )
    with pytest.raises(StableRecurringIdentityError, match="immutable IDs"):
        bridge.preview_recurring_change_from_excel(
            path, baseline_date=BASELINE,
            change=proposed(new_start_time=time(12, 15)),
            expected_entry=BASE, provider_max_daily_slots=CAPACITIES,
            require_stable_identity=True,
        )


def test_persistent_ready_preview_accepts_verified_stable_id(tmp_path, monkeypatch):
    from dataclasses import replace
    path = fixture_workbook(tmp_path, monkeypatch)
    stable_id = "assignment:STABLE0000123456"
    stable_entry = replace(BASE, base_entry_id=stable_id)
    monkeypatch.setattr(
        bridge, "bind_stable_assignment_ids",
        lambda path, entries: (stable_entry, OTHER),
    )
    change = RecurringTherapyChange(
        stable_id, "SYN-P", MOVE_DAY, new_start_time=time(12, 15),
    )
    preview = bridge.preview_recurring_change_from_excel(
        path, baseline_date=BASELINE, change=change,
        expected_entry=stable_entry, provider_max_daily_slots=CAPACITIES,
        require_stable_identity=True,
    )
    assert preview.proposed.change.base_entry_id == stable_id
