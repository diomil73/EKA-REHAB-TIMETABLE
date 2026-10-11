from __future__ import annotations

import importlib.util
from datetime import date, time
from pathlib import Path
from types import SimpleNamespace

from rehab_core.scheduling_candidates import TherapistScheduleCandidate
from rehab_excel.patient_scheduling_flow import TherapistChoice
from rehab_excel.patient_treatment_plan import (
    PatientTreatmentPlan,
    TreatmentSchedulingOptions,
)


MODULE_PATH = Path(__file__).resolve().parents[1] / "Python" / "tools" / "patient_scheduling_bridge.py"
spec = importlib.util.spec_from_file_location("patient_scheduling_bridge_tool", MODULE_PATH)
assert spec is not None and spec.loader is not None
bridge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bridge)


DAY = date(2026, 10, 12)


def test_plan_payload_is_vba_friendly_and_keeps_workload_label(monkeypatch):
    candidate = TherapistScheduleCandidate(
        therapist_id="T1",
        display_name="Γαύρας",
        active_sessions=3,
        active_timeslots=3,
        max_daily_timeslots=6,
        available_slots=(time(9, 0), time(10, 0)),
    )
    plan = PatientTreatmentPlan(
        patient_id="P001",
        target_date=DAY,
        treatments=(
            TreatmentSchedulingOptions(
                treatment="ΦΘ",
                requires_therapist=True,
                therapist_candidates=(candidate,),
            ),
        ),
    )
    monkeypatch.setattr(bridge, "build_patient_treatment_plan", lambda *args, **kwargs: plan)
    monkeypatch.setattr(
        bridge,
        "therapist_choices",
        lambda *args, **kwargs: (
            TherapistChoice("T1", "Γαύρας 3", (time(9, 0), time(10, 0))),
        ),
    )

    payload = bridge._plan_payload(
        SimpleNamespace(
            workbook=Path("dummy.xlsm"),
            patient_id="P001",
            date="2026-10-12",
            treatment=["ΦΘ"],
            robotic_capable=[],
        )
    )

    assert payload["ok"] is True
    assert payload["patient_id"] == "P001"
    therapist = payload["treatments"][0]["therapists"][0]
    assert therapist["label"] == "Γαύρας 3"
    assert therapist["available_slots"] == ["09:00", "10:00"]


def test_validate_payload_returns_current_revalidation_result(monkeypatch):
    monkeypatch.setattr(
        bridge,
        "refresh_and_validate_assignment",
        lambda *args, **kwargs: SimpleNamespace(valid=False, message="slot changed"),
    )

    payload = bridge._validate_payload(
        SimpleNamespace(
            workbook=Path("dummy.xlsm"),
            patient_id="P001",
            date="2026-10-12",
            treatment=["ΦΘ"],
            robotic_capable=[],
            assignment_treatment="ΦΘ",
            therapist_id="T1",
            start_time="10:00",
        )
    )

    assert payload == {"ok": True, "valid": False, "message": "slot changed"}
