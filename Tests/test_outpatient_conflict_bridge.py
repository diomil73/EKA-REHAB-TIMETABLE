from datetime import time

import rehab_excel.outpatient_schedule_bridge as bridge
from rehab_core.recurring_conflicts import RecurringTherapistConflict
from rehab_excel.outpatient_schedule_registration import (
    OutpatientTherapistDoubleBookingError,
)


def _payload(*, allow=False):
    return {
        "source_path": "book.xlsm",
        "preview_dir": ".",
        "overwrite": True,
        "values": {
            "patient_id": "P2",
            "treatment": "ΦΘ",
            "time": "09:15",
            "days": "Δ-Π",
            "therapist": "Χρήστου",
            "allow_double_booking": allow,
        },
    }


def test_bridge_returns_popup_ready_double_booking_details(monkeypatch):
    conflict = RecurringTherapistConflict(
        therapist_id="Χρήστου",
        start_time=time(9, 15),
        existing_patient_id="P1",
        existing_patient_name="ΔΙΑΤΣΙΝΤΟΣ",
        new_patient_id="P2",
        new_patient_name="ΠΑΠΑΓΕΩΡΓΙΟΥ",
        existing_day_pattern="Δ-Τε",
        new_day_pattern="Δ-Π",
        existing_base_entry_id="planner:54:ΦΘ",
    )

    def fail_with_conflict(*args, **kwargs):
        raise OutpatientTherapistDoubleBookingError(conflict)

    monkeypatch.setattr(bridge, "create_outpatient_schedule_preview", fail_with_conflict)

    result = bridge.run_outpatient_schedule_bridge(
        _payload(), fingerprint_service=lambda path: "abc"
    )

    assert result["ok"] is False
    assert result["conflict_type"] == "therapist_double_booking"
    assert result["therapist"] == "Χρήστου"
    assert result["time"] == "09:15"
    assert result["existing_patient_name"] == "ΔΙΑΤΣΙΝΤΟΣ"
    assert result["new_patient_name"] == "ΠΑΠΑΓΕΩΡΓΙΟΥ"


def test_bridge_passes_explicit_override_only_when_requested(monkeypatch):
    captured = {}

    class Report:
        source_path = "book.xlsm"
        output_path = "preview.xlsm"
        base_entry_id = "outpatient:2"
        excel_row = 2
        source_unchanged = True
        verified_in_output = True
        created = True

    def fake_create(*args, **kwargs):
        captured.update(kwargs)
        return Report()

    monkeypatch.setattr(bridge, "create_outpatient_schedule_preview", fake_create)
    monkeypatch.setattr(bridge, "asdict", lambda report: report.__dict__.copy())

    result = bridge.run_outpatient_schedule_bridge(
        _payload(allow=True), fingerprint_service=lambda path: "abc"
    )

    assert result["ok"] is True
    assert captured["allow_therapist_double_booking"] is True
