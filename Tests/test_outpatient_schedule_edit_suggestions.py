from types import SimpleNamespace

from rehab_excel.outpatient_schedule_suggestion_bridge import (
    run_outpatient_schedule_suggestion_bridge,
)


def test_edit_suggestions_ignore_current_authoritative_row(monkeypatch, tmp_path):
    source = tmp_path / "source.xlsm"
    source.write_bytes(b"placeholder")

    entries = (
        SimpleNamespace(base_entry_id="outpatient:4"),
        SimpleNamespace(base_entry_id="planner:P-IN:PT"),
    )
    captured = {}

    monkeypatch.setattr(
        "rehab_excel.outpatient_schedule_suggestion_bridge.read_settings",
        lambda _: SimpleNamespace(
            therapist_names=("T1", "T2"),
            standard_timeslots=(),
            day_patterns=(),
        ),
    )
    monkeypatch.setattr(
        "rehab_excel.outpatient_schedule_suggestion_bridge.read_unified_base_schedule",
        lambda _: entries,
    )

    def fake_suggest(**kwargs):
        captured["entries"] = tuple(kwargs["existing_entries"])
        return ()

    monkeypatch.setattr(
        "rehab_excel.outpatient_schedule_suggestion_bridge.suggest_outpatient_schedule_slots",
        fake_suggest,
    )

    response = run_outpatient_schedule_suggestion_bridge(
        {
            "source_path": str(source),
            "values": {
                "mode": 1,
                "therapist": "T1",
                "time": "09:00",
                "days": "Δε-Τε",
                "target_base_entry_id": "outpatient:4",
                "limit": 5,
            },
        }
    )

    assert response["ok"] is True
    assert [item.base_entry_id for item in captured["entries"]] == ["planner:P-IN:PT"]
