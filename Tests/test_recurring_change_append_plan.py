"""Synthetic-only preflight for the immutable, read-only event append plan."""
from datetime import date, time
from dataclasses import replace

import pytest

from rehab_core.models import BaseScheduleEntry
from rehab_core.recurring_therapy_changes import RecurringTherapyChange
from rehab_excel import recurring_change_append_plan as append
from rehab_excel.recurring_change_preview import RecurringExcelPreview

BASELINE = date(2026, 11, 1)
ASSIGNMENT = "assignment:SYNTHETIC123456"
BASE = BaseScheduleEntry(
    ASSIGNMENT, "SYN-P", "ΦΘ", time(9, 15), "Δε", "T-OLD",
)
OLD = RecurringTherapyChange(
    ASSIGNMENT, "SYN-P", date(2026, 11, 16),
    new_therapist_id="T-NEW",
)


def proposal(day=date(2026, 12, 1)):
    return RecurringTherapyChange(
        ASSIGNMENT, "SYN-P", day,
        new_start_time=time(10, 15), reason="Synthetic change",
    )


def prepared(tmp_path, monkeypatch, *, history=(), ids=None):
    source = tmp_path / "SYNTHETIC_DEMO.xlsm"
    source.write_bytes(b"THIS IS NOT A CLINICAL FILE")
    monkeypatch.setattr(append, "read_effective_base_entries",
                        lambda path: (replace(BASE, base_entry_id="weekly:2"),))
    monkeypatch.setattr(append, "bind_stable_assignment_ids",
                        lambda path, entries: (BASE,))
    monkeypatch.setattr(append, "_event_ids_and_shape",
                        lambda path: (set(ids if ids is not None else
                                          (f"change:EXISTINGEVENT0000{i}" for i in range(len(history)))), True))
    monkeypatch.setattr(append, "read_recurring_change_events",
                        lambda path, entries, baseline_date: tuple(history))

    def fake_preview(path, **kwargs):
        assert kwargs["require_stable_identity"] is True
        assert kwargs["current_changes"] == tuple(history)
        return RecurringExcelPreview(
            workbook_sha256=append.file_digest(path),
            base_entry_count=1,
            proposed=object(),
        )
    monkeypatch.setattr(append, "preview_recurring_change_from_excel", fake_preview)
    return source


def test_append_plan_is_readonly_and_source_hash_bound(tmp_path, monkeypatch):
    source = prepared(tmp_path, monkeypatch, history=(OLD,))
    before = source.read_bytes()
    result = append.plan_recurring_event_append(
        source, event_id="change:NEWOPAQUEEVENT000001",
        change=proposal(), expected_entry=BASE,
        baseline_date=BASELINE, provider_max_daily_slots={"T-OLD": 4, "T-NEW": 6},
    )
    assert result.source_sha256 == append.file_digest(source)
    assert result.event_count_before == 1
    assert result.base_entry_count == 1
    assert result.row_values[:4] == (
        "change:NEWOPAQUEEVENT000001", ASSIGNMENT, "SYN-P", "2026-12-01"
    )
    assert source.read_bytes() == before
    append.verify_unchanged_source(source, result)
    source.write_bytes(b"EXTERNAL CHANGE")
    with pytest.raises(append.RecurringAppendPlanError, match="Source changed"):
        append.verify_unchanged_source(source, result)


def test_rejects_backdated_append_for_same_assignment(tmp_path, monkeypatch):
    source = prepared(tmp_path, monkeypatch, history=(OLD,))
    with pytest.raises(append.RecurringAppendPlanError, match="later than prior"):
        append.plan_recurring_event_append(
            source, event_id="change:NEWOPAQUEEVENT000001",
            change=proposal(date(2026, 11, 12)), expected_entry=BASE,
            baseline_date=BASELINE, provider_max_daily_slots={"T-OLD": 4},
        )


def test_rejects_duplicate_event_id(tmp_path, monkeypatch):
    existing_id = "change:EXISTINGEVENT00000"
    source = prepared(tmp_path, monkeypatch, history=(OLD,), ids={existing_id})
    with pytest.raises(append.RecurringAppendPlanError, match="Duplicate event ID"):
        append.plan_recurring_event_append(
            source, event_id=existing_id, change=proposal(),
            expected_entry=BASE, baseline_date=BASELINE,
            provider_max_daily_slots={"T-OLD": 4},
        )


def test_rejects_legacy_assignment_id_before_reading_excel(tmp_path):
    source = tmp_path / "SYNTHETIC_DEMO.xlsm"
    source.write_bytes(b"SYNTHETIC")
    with pytest.raises(append.RecurringAppendPlanError, match="Immutable"):
        append.plan_recurring_event_append(
            source, event_id="change:NEWOPAQUEEVENT000001",
            change=RecurringTherapyChange(
                "weekly:2", "SYN-P", date(2026, 12, 1),
                new_start_time=time(11, 15),
            ),
            expected_entry=BASE, baseline_date=BASELINE,
            provider_max_daily_slots={"T-OLD": 4},
        )


def test_rejects_edited_assignment_fingerprint(tmp_path, monkeypatch):
    source = prepared(tmp_path, monkeypatch)
    with pytest.raises(append.RecurringAppendPlanError, match="Stale assignment"):
        append.plan_recurring_event_append(
            source, event_id="change:NEWOPAQUEEVENT000001",
            change=proposal(),
            expected_entry=replace(BASE, start_time=time(10, 15)),
            baseline_date=BASELINE, provider_max_daily_slots={"T-OLD": 4},
        )


def test_detects_concurrent_source_change_while_reading(tmp_path, monkeypatch):
    source = prepared(tmp_path, monkeypatch)
    def mutate(path, entries):
        source.write_bytes(b"CHANGED CONCURRENTLY")
        return (BASE,)
    monkeypatch.setattr(append, "bind_stable_assignment_ids", mutate)
    with pytest.raises(append.RecurringAppendPlanError, match="changed during"):
        append.plan_recurring_event_append(
            source, event_id="change:NEWOPAQUEEVENT000001",
            change=proposal(), expected_entry=BASE,
            baseline_date=BASELINE, provider_max_daily_slots={"T-OLD": 4},
        )
