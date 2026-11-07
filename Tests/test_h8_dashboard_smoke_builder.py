from pathlib import Path
from types import SimpleNamespace

import pytest

from tools import prepare_h8_dashboard_smoke as builder


def test_h8_build_runs_h7_then_new_refresh_installer(monkeypatch, tmp_path):
    source = tmp_path / "AZ.xlsm"
    output = tmp_path / "BJ.xlsm"
    source.write_bytes(b"frozen-source")
    calls = []

    def fake_h7(src, dst):
        calls.append(("h7", src, dst))
        dst.write_bytes(src.read_bytes())

    def fake_subprocess(args, check=False):
        calls.append(("h8", tuple(args), check))
        return SimpleNamespace(returncode=0)

    monkeypatch.setattr(builder, "prepare_h7_integrated_ui", fake_h7)
    monkeypatch.setattr(builder.subprocess, "run", fake_subprocess)
    builder.prepare_h8_dashboard_smoke(source, output)
    assert [x[0] for x in calls] == ["h7", "h8"]
    assert calls[1][1][1].endswith("install_h8_single_message_refresh.py")
    assert source.read_bytes() == output.read_bytes()


def test_h8_build_refuses_existing_destination_via_h7_safety_gate(tmp_path, monkeypatch):
    source = tmp_path / "AZ.xlsm"
    output = tmp_path / "frozen.xlsm"
    source.write_bytes(b"source")
    output.write_bytes(b"already-frozen")
    monkeypatch.setattr(
        builder.subprocess, "run",
        lambda *args, **kwargs: pytest.fail("Must not invoke H8 stage on existing output"),
    )
    with pytest.raises(builder.H8DashboardPreparationError, match="already exists"):
        builder.prepare_h8_dashboard_smoke(source, output)
    assert output.read_bytes() == b"already-frozen"


def test_h8_installer_failure_freezes_new_destination(monkeypatch, tmp_path):
    source = tmp_path / "AZ.xlsm"
    output = tmp_path / "BJ.xlsm"
    source.write_bytes(b"trusted")
    monkeypatch.setattr(
        builder, "prepare_h7_integrated_ui",
        lambda src, dst: dst.write_bytes(src.read_bytes()),
    )
    monkeypatch.setattr(
        builder.subprocess, "run", lambda *args, **kwargs: SimpleNamespace(returncode=2),
    )
    with pytest.raises(builder.H8DashboardPreparationError, match="freeze destination"):
        builder.prepare_h8_dashboard_smoke(source, output)
    assert output.exists()
    assert source.read_bytes() == b"trusted"
