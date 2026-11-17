"""Synthetic-only private real-data DEMO entry-point regression."""
from pathlib import Path
import json
from types import SimpleNamespace

import pytest

from rehab_core.models import PatientType
from tools import prepare_real_data_practice_demo as cli


@pytest.fixture(autouse=True)
def _mock_private_special_source_for_fake_bytes(monkeypatch):
    # These tests use intentionally non-XLSX fake source bytes. The real
    # original-F/U/V workbook audit is exercised independently in
    # Tests/test_trial_special_fields_audit.py, never bypassed in production.
    monkeypatch.setattr(
        cli, "inspect_real_manual_special_fields",
        lambda path: SimpleNamespace(unresolved_fields=0, orphan_rows=0),
    )


def _approvals(tmp_path, records=None):
    source = tmp_path / "PRIVATE_APPROVALS.json"
    if records is None:
        records = [
            {
                "clinic": "DEMO CLINIC A",
                "source_row": 2,
                "patient_type": "INPATIENT",
                "infectious": False,
                "room": "A1",
                "status": "Ενεργός",
                "responsible_doctor": None,
            },
            {
                "clinic": "DEMO CLINIC B",
                "source_row": 3,
                "patient_type": "OUTPATIENT",
                "infectious": False,
                "room": None,
            },
        ]
    source.write_text(json.dumps(
        {"schema": cli.APPROVAL_SCHEMA, "records": records},
        ensure_ascii=False,
    ), encoding="utf-8")
    return source


def _inputs(tmp_path):
    paths = [
        tmp_path / "PRIVATE_MANUAL.xlsx",
        tmp_path / "PRIVATE_REVIEW.xlsx",
        tmp_path / "PRIVATE_EMPTY_SHELL.xlsm",
        _approvals(tmp_path),
    ]
    for path in paths[:3]:
        path.write_bytes(b"synthetic test file only")
    return paths


def test_private_clinical_approval_parser_requires_explicit_types_and_flags(tmp_path):
    rows = cli.read_private_clinical_approvals(_approvals(tmp_path))
    assert len(rows) == 2
    assert rows[("DEMO CLINIC A", 2)].patient_type == PatientType.INPATIENT
    assert rows[("DEMO CLINIC B", 3)].patient_type == PatientType.OUTPATIENT
    assert rows[("DEMO CLINIC B", 3)].infectious is False
    assert rows[("DEMO CLINIC B", 3)].room is None


@pytest.mark.parametrize("mutation", [
    {"patient_type": "UNKNOWN"},
    {"infectious": "No"},
    {"infectious": None},
    {"source_row": True},
    {"source_row": 1},
    {"clinic": ""},
    {"clinic": "DEMO CLINIC A "},
    {"hospital_mrn": 1234},
    {"new_unknown_field": "guess"},
])
def test_private_approvals_fail_closed_on_invalid_clinical_record(tmp_path, mutation):
    record = {
        "clinic": "DEMO CLINIC A", "source_row": 2,
        "patient_type": "INPATIENT", "infectious": False,
    }
    record.update(mutation)
    with pytest.raises(cli.RealDemoCommandError):
        cli.read_private_clinical_approvals(_approvals(tmp_path, [record]))


def test_private_approvals_reject_casefold_duplicate_source_identity(tmp_path):
    rows = [
        {"clinic": "DEMO A", "source_row": 2, "patient_type": "INPATIENT", "infectious": False},
        {"clinic": "demo a", "source_row": 2, "patient_type": "INPATIENT", "infectious": False},
    ]
    with pytest.raises(cli.RealDemoCommandError, match="Duplicate"):
        cli.read_private_clinical_approvals(_approvals(tmp_path, rows))


def test_json_duplicate_keys_are_refused_without_showing_record(tmp_path):
    path = tmp_path / "APPROVALS.json"
    path.write_text(
        '{"schema":"diomil-private-clinical-approvals-v1",'
        '"records":[{"clinic":"DEMO","clinic":"PRIVATE","source_row":2,'
        '"patient_type":"INPATIENT","infectious":false}]}',
        encoding="utf-8",
    )
    with pytest.raises(cli.RealDemoCommandError, match="Duplicate JSON"):
        cli.read_private_clinical_approvals(path)


def test_preflight_only_writes_no_native_workbook(tmp_path, monkeypatch):
    p, r, shell, approval = _inputs(tmp_path)
    calls = []
    monkeypatch.setattr(cli, "assert_clean_native_demo_shell", lambda path: calls.append("shell"))
    monkeypatch.setattr(cli, "prepare_trial_import_plan", lambda **kwargs: (
        calls.append("plan") or SimpleNamespace(verify_sources=lambda *args, **kwargs: calls.append("hash"))
    ))
    monkeypatch.setattr(cli, "project_trial_demo_sources", lambda *args, **kwargs: (
        calls.append("project") or SimpleNamespace(patients=2, weekday_slot_count=12)
    ))
    monkeypatch.setattr(cli, "seed_trial_demo_workbook", lambda *args, **kwargs: (
        pytest.fail("preflight may never write a native DEMO")
    ))
    result = cli.run_private_trial(
        programme=p, review=r, shell=shell, approvals=approval,
    )
    assert result == (2, 12, False)
    assert calls == ["shell", "plan", "project", "hash"]


def _fake_schedule(tmp_path, monkeypatch):
    path = tmp_path / "PRIVATE_PROVIDER_SCHEDULE.json"
    path.write_bytes(b"fictional mock schedule")
    policy = object()
    monkeypatch.setattr(cli, "read_private_provider_schedule", lambda *_: policy)
    monkeypatch.setattr(cli, "merge_with_approved_provider_schedule", lambda projection, *a, **kw: projection)
    return path, policy


def test_native_build_is_explicit_and_injects_only_approved_source(monkeypatch, tmp_path):
    p, r, shell, approval = _inputs(tmp_path)
    destination = tmp_path / "PRIVATE_DEMO_NEW.xlsm"
    schedule, policy = _fake_schedule(tmp_path, monkeypatch)
    calls = []
    class FakePlan:
        def verify_sources(self, *args, **kwargs):
            calls.append("hash")
    monkeypatch.setattr(cli.sys, "platform", "win32")
    monkeypatch.setattr(cli, "assert_clean_native_demo_shell", lambda x: calls.append("shell"))
    monkeypatch.setattr(cli, "prepare_trial_import_plan", lambda **kw: (
        calls.append("plan") or FakePlan()
    ))
    monkeypatch.setattr(cli, "project_trial_demo_sources", lambda *a, **kw: (
        calls.append("project") or SimpleNamespace(patients=2, weekday_slot_count=12)
    ))
    def fake_seed(plan, **kw):
        assert kw["programme"] == p
        assert kw["review"] == r
        assert kw["intake"] == shell
        assert kw["output"] == destination
        assert len(kw["clinical_approvals"]) == 2
        assert kw["provider_schedule"] is policy
        calls.append("native")
        return SimpleNamespace(
            patient_count=2, weekday_slot_count=12,
            source_unchanged=True, verified_native_readback=True,
        )
    monkeypatch.setattr(cli, "seed_trial_demo_workbook", fake_seed)
    monkeypatch.setattr(cli, "project_private_outpatient_source", lambda path: object())
    monkeypatch.setattr(cli, "merge_trial_demo_outpatients", lambda projection, outpatient: projection)
    outpatients = tmp_path / "PRIVATE_APPROVED_OUTPATIENTS.xlsx"
    outpatients.write_bytes(b"fake test only")
    assert cli.run_private_trial(
        programme=p, review=r, shell=shell, approvals=approval,
        output=destination, build=True, outpatients=outpatients, provider_schedule=schedule,
    ) == (2, 12, True)
    assert calls == ["shell", "plan", "project", "hash", "native"]


def test_output_never_overwrites_and_must_stay_outside_git(tmp_path, monkeypatch):
    p, r, shell, approval = _inputs(tmp_path)
    monkeypatch.setattr(cli.sys, "platform", "win32")
    existing = tmp_path / "PRIVATE_DEMO_NEW.xlsm"
    existing.write_bytes(b"do not overwrite")
    outpatient = tmp_path / "APPROVED_OUTPATIENTS.xlsx"
    outpatient.write_bytes(b"fictional")
    with pytest.raises(cli.RealDemoCommandError, match="Unsafe"):
        cli.validate_private_paths(
            programme=p, review=r, shell=shell, approvals=approval,
            output=existing, build=True, outpatients=outpatient,
        )
    git = tmp_path / "repo"
    git.mkdir()
    (git / ".git").mkdir()
    embedded = git / "PRIVATE_DEMO_OTHER.xlsm"
    with pytest.raises(cli.RealDemoCommandError, match="Unsafe"):
        cli.validate_private_paths(
            programme=p, review=r, shell=shell, approvals=approval,
            output=embedded, build=True, outpatients=outpatient,
        )


def test_git_checkout_source_rejected_even_during_read_only_preflight(tmp_path):
    p, r, shell, approval = _inputs(tmp_path)
    (tmp_path / ".git").mkdir()
    with pytest.raises(cli.RealDemoCommandError, match="outside Git"):
        cli.validate_private_paths(
            programme=p, review=r, shell=shell, approvals=approval,
            output=None, build=False,
        )


def test_command_does_not_leak_private_exception_message(tmp_path, monkeypatch, capsys):
    p, r, shell, approval = _inputs(tmp_path)
    monkeypatch.setattr(cli, "run_private_trial", lambda **kwargs: (
        (_ for _ in ()).throw(ValueError("PRIVATE NAME AND MRN"))
    ))
    result = cli.main([
        "--programme", str(p), "--review", str(r),
        "--shell", str(shell), "--clinical-approvals", str(approval),
    ])
    output = capsys.readouterr().out
    assert result == 2
    assert "SAFETY STOP" in output
    assert "PRIVATE" not in output
    for path in (p, r, shell, approval):
        assert str(path) not in output


def test_preflight_output_argument_without_build_is_rejected(tmp_path):
    p, r, shell, approval = _inputs(tmp_path)
    with pytest.raises(cli.RealDemoCommandError, match="only allowed"):
        cli.validate_private_paths(
            programme=p, review=r, shell=shell, approvals=approval,
            output=tmp_path / "PRIVATE_DEMO_NEW.xlsm", build=False,
        )



def test_changed_private_clinical_approval_bytes_stop_before_native_build(
    tmp_path, monkeypatch,
):
    p, r, shell, approval = _inputs(tmp_path)
    monkeypatch.setattr(cli, "assert_clean_native_demo_shell", lambda path: None)
    class FakePlan:
        def verify_sources(self, *args, **kwargs):
            pass
    def change_approval(**kwargs):
        approval.write_bytes(approval.read_bytes() + b" ")
        return FakePlan()
    monkeypatch.setattr(cli, "prepare_trial_import_plan", change_approval)
    monkeypatch.setattr(cli, "project_trial_demo_sources", lambda *a, **kw: (
        SimpleNamespace(patients=2, weekday_slot_count=12)
    ))
    monkeypatch.setattr(cli, "seed_trial_demo_workbook", lambda *a, **kw: (
        pytest.fail("changed clinical approval must never enter native seed")
    ))
    with pytest.raises(cli.RealDemoCommandError, match="changed"):
        cli.run_private_trial(
            programme=p, review=r, shell=shell, approvals=approval,
        )


def test_changed_private_approvals_during_native_build_not_accepted(
    tmp_path, monkeypatch,
):
    p, r, shell, approval = _inputs(tmp_path)
    monkeypatch.setattr(cli.sys, "platform", "win32")
    monkeypatch.setattr(cli, "assert_clean_native_demo_shell", lambda path: None)
    monkeypatch.setattr(cli, "prepare_trial_import_plan", lambda **kw: (
        SimpleNamespace(verify_sources=lambda *a, **kw: None)
    ))
    monkeypatch.setattr(cli, "project_trial_demo_sources", lambda *a, **kw: (
        SimpleNamespace(patients=2, weekday_slot_count=12)
    ))
    def unsafe_save(*a, **kw):
        approval.write_bytes(approval.read_bytes() + b" ")
        return SimpleNamespace(
            patient_count=2, weekday_slot_count=12,
            source_unchanged=True, verified_native_readback=True,
        )
    monkeypatch.setattr(cli, "seed_trial_demo_workbook", unsafe_save)
    outpatient = tmp_path / "APPROVED_OUTPATIENTS.xlsx"
    outpatient.write_bytes(b"synthetic only")
    schedule, _ = _fake_schedule(tmp_path, monkeypatch)
    monkeypatch.setattr(cli, "project_private_outpatient_source", lambda path: object())
    monkeypatch.setattr(cli, "merge_trial_demo_outpatients", lambda projection, outpatient: projection)
    with pytest.raises(cli.RealDemoCommandError, match="changed"):
        cli.run_private_trial(
            programme=p, review=r, shell=shell, approvals=approval,
            output=tmp_path / "DIOMIL_DEMO_PRIVATE_B.xlsm",
            build=True, outpatients=outpatient, provider_schedule=schedule,
        )


def test_special_fuv_approval_gate_blocks_even_preflight_before_any_native_action(
    tmp_path, monkeypatch,
):
    p, r, shell, approval = _inputs(tmp_path)
    monkeypatch.setattr(
        cli, "inspect_real_manual_special_fields",
        lambda path: SimpleNamespace(unresolved_fields=1, orphan_rows=0),
    )
    monkeypatch.setattr(
        cli, "assert_clean_native_demo_shell",
        lambda *_: pytest.fail("no shell processing if original F/U/V has holds"),
    )
    monkeypatch.setattr(
        cli, "seed_trial_demo_workbook",
        lambda *a, **kw: pytest.fail("no native seeding of unreviewed special rows"),
    )
    with pytest.raises(cli.RealDemoCommandError, match="F/U/V"):
        cli.run_private_trial(
            programme=p, review=r, shell=shell, approvals=approval,
        )


def test_special_formula_only_source_blocks_native_build_before_approval_read(
    tmp_path, monkeypatch,
):
    p, r, shell, approval = _inputs(tmp_path)
    monkeypatch.setattr(cli.sys, "platform", "win32")
    monkeypatch.setattr(
        cli, "inspect_real_manual_special_fields",
        lambda path: SimpleNamespace(unresolved_fields=1, orphan_rows=1),
    )
    monkeypatch.setattr(
        cli, "read_private_clinical_approvals",
        lambda *_: pytest.fail("no clinical application on held original formulas"),
    )
    # The independent source exists, so validation reaches the F/U/V gate.
    # It must stop before any clinical approval or Excel mutation occurs.
    outpatient = tmp_path / "SEPARATE_OUTPATIENTS.xlsx"
    outpatient.write_bytes(b"synthetic standalone outpatient source")
    with pytest.raises(cli.RealDemoCommandError, match="F/U/V"):
        cli.run_private_trial(
            programme=p, review=r, shell=shell, approvals=approval,
            build=True, output=tmp_path / "DIOMIL_DEMO_HELD.xlsm",
            outpatients=outpatient,
        )


def test_private_native_build_requires_a_distinct_approved_outpatient_source(tmp_path):
    p, r, shell, approval = _inputs(tmp_path)
    with pytest.raises(cli.RealDemoCommandError, match="outpatients"):
        cli.validate_private_paths(
            programme=p, review=r, shell=shell, approvals=approval,
            output=tmp_path / "DIOMIL_DEMO_REQUIRED.xlsm", build=True,
        )


def test_outpatient_source_cannot_reuse_inpatient_input_file(tmp_path):
    p, r, shell, approval = _inputs(tmp_path)
    with pytest.raises(cli.RealDemoCommandError, match="distinct"):
        cli.validate_private_paths(
            programme=p, review=r, shell=shell, approvals=approval,
            output=None, build=False, outpatients=p,
        )


def test_private_outpatient_projection_not_called_when_source_absent(tmp_path, monkeypatch):
    p, r, shell, approval = _inputs(tmp_path)
    monkeypatch.setattr(cli, "assert_clean_native_demo_shell", lambda *_: None)
    monkeypatch.setattr(cli, "prepare_trial_import_plan", lambda **kw: (
        SimpleNamespace(verify_sources=lambda *a, **kw: None)
    ))
    monkeypatch.setattr(cli, "project_trial_demo_sources", lambda *a, **kw: (
        SimpleNamespace(patients=1, weekday_slot_count=5)
    ))
    monkeypatch.setattr(cli, "project_private_outpatient_source", lambda *_: (
        pytest.fail("No separate source means no outpatient inference")
    ))
    assert cli.run_private_trial(
        programme=p, review=r, shell=shell, approvals=approval,
    ) == (1, 5, False)


def test_native_build_without_provider_schedule_is_blocked_before_projection(tmp_path, monkeypatch):
    p, r, shell, approval = _inputs(tmp_path)
    outpatient = tmp_path / "PRIVATE_OUTPATIENTS.xlsx"
    outpatient.write_bytes(b"fictional")
    monkeypatch.setattr(cli.sys, "platform", "win32")
    monkeypatch.setattr(cli, "assert_clean_native_demo_shell", lambda *_: pytest.fail("missing schedule must stop"))
    with pytest.raises(cli.RealDemoCommandError, match="approved provider schedule"):
        cli.run_private_trial(
            programme=p, review=r, shell=shell, approvals=approval, outpatients=outpatient,
            build=True, output=tmp_path / "PRIVATE_DEMO_NEW.xlsm",
        )
