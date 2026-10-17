from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
from datetime import date, time
from pathlib import Path
from types import SimpleNamespace

from rehab_excel.automatic_patient_schedule import (
    AutomaticProposedAssignment,
    AutomaticScheduleProposal,
    AutomaticUnresolvedTreatment,
)


MODULE_PATH = Path(__file__).resolve().parents[1] / "Python" / "tools" / "patient_scheduling_bridge.py"
spec = importlib.util.spec_from_file_location("patient_scheduling_bridge_tool", MODULE_PATH)
assert spec is not None and spec.loader is not None
bridge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bridge)


DAY = date(2026, 10, 12)


def test_bridge_bootstraps_python_root_when_launched_directly(tmp_path):
    env = dict(os.environ)
    env.pop("PYTHONPATH", None)
    result = subprocess.run(
        [sys.executable, str(MODULE_PATH), "--help"],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        timeout=15,
    )
    assert result.returncode == 0, result.stderr
    assert "Patient scheduling runtime bridge" in result.stdout


def test_bridge_exposes_only_joint_auto_commands():
    parser = bridge.build_parser()
    action = next(action for action in parser._actions if action.dest == "command")
    assert tuple(action.choices) == ("auto-propose", "auto-propose-vba")


def test_auto_payload_exposes_assignments_and_unresolved(monkeypatch):
    proposal = AutomaticScheduleProposal(
        patient_id="P001",
        anchor_date=DAY,
        assignments=(
            AutomaticProposedAssignment(
                treatment="ΦΘ",
                day_pattern="Τρ-Πε",
                therapist_id="T2",
                therapist_name="Μηλιδάκης",
                start_time=time(8, 30),
                workload_count=2,
            ),
        ),
        unresolved=(
            AutomaticUnresolvedTreatment(
                treatment="Εργο",
                day_pattern="Καθ/να",
                reason="unsupported",
            ),
        ),
        total_workload=2,
    )
    monkeypatch.setattr(
        bridge,
        "propose_automatic_patient_schedule",
        lambda *args, **kwargs: proposal,
    )

    payload = bridge._auto_payload(
        SimpleNamespace(
            workbook=Path("dummy.xlsm"),
            patient_id="P001",
            date="2026-10-12",
            request=["ΦΘ|Τρ-Πε", "Εργο|Καθ/να"],
            robotic_capable=[],
        )
    )

    assert payload["complete"] is False
    assert payload["total_workload"] == 2
    assert payload["assignments"][0]["therapist_name"] == "Μηλιδάκης"
    assert payload["assignments"][0]["start_time"] == "08:30"
    assert payload["unresolved"][0] == {
        "treatment": "Εργο",
        "day_pattern": "Καθ/να",
        "reason": "unsupported",
    }


def test_auto_vba_lines_are_tab_safe_and_explicit():
    lines = bridge._auto_vba_lines(
        {
            "patient_id": "P001",
            "date": "2026-10-12",
            "complete": False,
            "total_workload": 2,
            "assignments": [
                {
                    "treatment": "ΦΘ",
                    "day_pattern": "Τρ-Πε",
                    "therapist_id": "T2",
                    "therapist_name": "Μηλιδάκης",
                    "start_time": "08:30",
                    "workload_count": 2,
                }
            ],
            "unresolved": [
                {
                    "treatment": "Εργο",
                    "day_pattern": "Καθ/να",
                    "reason": "not yet supported",
                }
            ],
        }
    )

    assert lines[0] == "OK\tP001\t2026-10-12\t0\t2"
    assert lines[1] == "ASSIGNMENT\tΦΘ\tΤρ-Πε\tT2\tΜηλιδάκης\t08:30\t2"
    assert lines[2] == "UNRESOLVED\tΕργο\tΚαθ/να\tnot yet supported"
