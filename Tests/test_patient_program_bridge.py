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


def test_program_bridge_emits_every_treatment_provider_and_existing_program(monkeypatch, capsys):
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
                        "treatment": "Εργο",
                        "provider_kind": "occupational_therapist",
                        "requires_provider": False,
                        "provider_choices": (),
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
    existing = type(
        "Entry",
        (),
        {
            "patient_id": "P001",
            "treatment": "Εργο",
            "day_pattern": "Τρ-Πε",
            "start_time": time(11, 0),
            "therapist_id": "Χειροκίνητο Όνομα",
        },
    )()
    monkeypatch.setattr(bridge, "build_patient_program_draft", lambda *a, **k: fake_draft)
    monkeypatch.setattr(bridge, "read_settings", lambda *a, **k: fake_settings)
    monkeypatch.setattr(bridge, "read_base_schedule", lambda *a, **k: [existing])

    bridge._emit_config(Path("dummy.xlsm"), "P001")
    output = capsys.readouterr().out

    assert "TREATMENT\t1\tΦΘ\tphysiotherapist\t1" in output
    assert "PROVIDER\tΦΘ\tΦυσιο 1" in output
    assert "TREATMENT\t2\tΕργο\toccupational_therapist\t0" in output
    assert "TREATMENT\t3\tΑνακλινόμενο\tnone\t0" in output
    assert "CURRENT\tΕργο\tΤρ-Πε\t11:00\tΧειροκίνητο Όνομα" in output
    assert "DAY_PATTERN\tΔε-Τε-Πα" in output
    assert "STANDARD_SLOT\t10:00" in output
    assert "RECLINED_SLOT\t09:30" in output


def test_parse_completion_request_preserves_manual_time_and_provider():
    request = bridge._parse_request("ΦΘ|Δε-Τε-Πα|10:30|Φυσιο 2")
    assert request.treatment == "ΦΘ"
    assert request.day_pattern == "Δε-Τε-Πα"
    assert request.start_time == time(10, 30)
    assert request.provider_id == "Φυσιο 2"


def test_parse_completion_request_leaves_missing_day_for_solver():
    request = bridge._parse_request("ΕΦΑ|||")
    assert request.treatment == "ΕΦΑ"
    assert request.day_pattern is None
    assert request.start_time is None
    assert request.provider_id is None


def test_completion_sends_missing_day_pattern_to_solver(monkeypatch, capsys):
    observed = {}
    assignment = type(
        "Assignment",
        (),
        {
            "treatment": "ΕΦΑ",
            "day_pattern": "Τρ-Πε",
            "start_time": time(10, 0),
            "therapist_id": "",
            "therapist_name": "",
            "workload_count": 0,
        },
    )()
    proposal = type("Proposal", (), {"assignments": (assignment,), "unresolved": ()})()

    def fake_solver(*args, requests, **kwargs):
        observed["requests"] = requests
        return proposal

    monkeypatch.setattr(bridge, "propose_automatic_patient_schedule", fake_solver)
    bridge._emit_completion(Path("dummy.xlsm"), "P001", ["ΕΦΑ|||"])
    output = capsys.readouterr().out

    assert observed["requests"][0].day_pattern is None
    assert "ASSIGNMENT\tΕΦΑ\tΤρ-Πε\t10:00\t\t\t0" in output


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
