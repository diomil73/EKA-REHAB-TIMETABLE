from datetime import time
from pathlib import Path
import subprocess
import sys

from tools import patient_program_bridge as bridge


def test_program_bridge_parser_accepts_config_command():
    parser = bridge.build_parser()
    args = parser.parse_args(
        ["config-vba", "dummy.xlsm", "--patient-id", "P001"]
    )
    assert args.command == "config-vba"
    assert args.patient_id == "P001"


def test_program_bridge_parser_accepts_completion_requests():
    parser = bridge.build_parser()
    args = parser.parse_args(
        [
            "complete-vba",
            "dummy.xlsm",
            "--patient-id",
            "P001",
            "--request",
            "ΦΘ|Δε-Τε-Πα|10:00|Φυσιο 1",
        ]
    )
    assert args.command == "complete-vba"
    assert args.request == ["ΦΘ|Δε-Τε-Πα|10:00|Φυσιο 1"]


def test_program_bridge_is_direct_launchable_without_import_error():
    script = Path(bridge.__file__).resolve()
    completed = subprocess.run(
        [sys.executable, str(script), "--help"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    assert completed.returncode == 0
    assert "Patient program editor runtime bridge" in completed.stdout


def test_program_bridge_emits_every_treatment_and_provider(monkeypatch, capsys):
    fake_draft = type(
        "Draft",
        (),
        {
            "patient_id": "P001",
            "treatments": (
                type(
                    "Spec",
                    (),
                    {
                        "treatment": "ΦΘ",
                        "provider_kind": "physiotherapist",
                        "requires_provider": True,
                        "provider_choices": ("Φυσιο 1", "Φυσιο 2"),
                    },
                )(),
                type(
                    "Spec",
                    (),
                    {
                        "treatment": "Ανακλινόμενο",
                        "provider_kind": "none",
                        "requires_provider": False,
                        "provider_choices": (),
                    },
                )(),
            ),
        },
    )()
    fake_settings = type(
        "Settings",
        (),
        {
            "day_patterns": ("Δε-Τε-Πα",),
            "standard_timeslots": (time(10, 0),),
            "reclined_timeslots": (time(9, 30),),
        },
    )()
    monkeypatch.setattr(bridge, "build_patient_program_draft", lambda *a, **k: fake_draft)
    monkeypatch.setattr(bridge, "read_settings", lambda *a, **k: fake_settings)

    bridge._emit_config(Path("dummy.xlsm"), "P001")
    output = capsys.readouterr().out

    assert "TREATMENT\t1\tΦΘ\tphysiotherapist\t1" in output
    assert "PROVIDER\tΦΘ\tΦυσιο 1" in output
    assert "TREATMENT\t2\tΑνακλινόμενο\tnone\t0" in output
    assert "DAY_PATTERN\tΔε-Τε-Πα" in output
    assert "STANDARD_SLOT\t10:00" in output
    assert "RECLINED_SLOT\t09:30" in output


def test_parse_completion_request_preserves_manual_time_and_provider():
    request = bridge._parse_request("ΦΘ|Δε-Τε-Πα|10:30|Φυσιο 2")
    assert request.treatment == "ΦΘ"
    assert request.day_pattern == "Δε-Τε-Πα"
    assert request.start_time == time(10, 30)
    assert request.provider_id == "Φυσιο 2"


def test_completion_marks_missing_day_pattern_as_unresolved(monkeypatch, capsys):
    bridge._emit_completion(Path("dummy.xlsm"), "P001", ["ΦΘ|||", "Ρομποτικό|||" ])
    output = capsys.readouterr().out
    assert output.count("UNRESOLVED") == 2
    assert "Χρειάζεται πρώτα επιλογή ημερών" in output


def test_completion_emits_solver_assignments(monkeypatch, capsys):
    assignment = type(
        "Assignment",
        (),
        {
            "treatment": "ΦΘ",
            "day_pattern": "Δε-Τε-Πα",
            "start_time": time(10, 0),
            "therapist_id": "T2",
            "therapist_name": "Φυσιο 2",
            "workload_count": 3,
        },
    )()
    proposal = type("Proposal", (), {"assignments": (assignment,), "unresolved": ()})()
    monkeypatch.setattr(
        bridge,
        "propose_automatic_patient_schedule",
        lambda *a, **k: proposal,
    )

    bridge._emit_completion(
        Path("dummy.xlsm"),
        "P001",
        ["ΦΘ|Δε-Τε-Πα|||"],
    )
    output = capsys.readouterr().out
    assert "ASSIGNMENT\tΦΘ\tΔε-Τε-Πα\t10:00\tT2\tΦυσιο 2\t3" in output
