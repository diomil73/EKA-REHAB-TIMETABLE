"""Tests for safe, aggregate-only native projection diagnostic CLI."""
from datetime import date
from types import SimpleNamespace

import Python.tools.diagnose_daily_input_projection_native as cli


def _result():
    native = SimpleNamespace(
        input_ledger_source_changes=3,
        input_ledger_active_rows=3,
        master_projected_rows=12,
        therapist_grid_cells_checked=126,
        therapist_occupied_cells=17,
        master_excel_version="16.0",
        therapist_excel_version="16.0",
    )
    return SimpleNamespace(
        service_date=date(2026, 11, 23),
        source_workbook_sha256="a" * 64,
        native_source_sheet_count=12,
        native_source_cells_checked=1450,
        patient_count=100,
        base_assignment_count=300,
        session_count=80,
        daily_change_count=3,
        replacement_count=1,
        therapist_assignment_count=77,
        native_readback=native,
    )


def test_cli_runs_readonly_preflight_and_prints_only_aggregate_results(monkeypatch, capsys):
    calls = []
    preflight = object()
    monkeypatch.setattr(cli.sys, "platform", "win32")
    monkeypatch.setattr(
        cli, "build_daily_input_projection_preflight",
        lambda source, **kwargs: calls.append(("preflight", source, kwargs)) or preflight,
    )
    monkeypatch.setattr(
        cli, "verify_daily_input_projection_preflight_native_readonly",
        lambda source, **kwargs: calls.append(("native", source, kwargs)) or _result(),
    )

    code = cli.main([
        "--source", "private.xlsm",
        "--service-date", "2026-11-23",
        "--baseline-date", "2026-01-05",
    ])

    output = capsys.readouterr().out
    assert code == 0
    assert "DAILY_INPUT PROJECTION NATIVE READBACK GREEN" in output
    assert "NON-AUTHORIZING READ-ONLY DIAGNOSTIC ONLY" in output
    assert "native_source_sheets_checked: 12" in output
    assert "native_source_populated_cells_checked: 1450" in output
    assert "patients: 100" in output
    assert "therapist_patient_cells_checked: 126" in output
    assert "private.xlsm" not in output
    assert calls[0][0] == "preflight"
    assert calls[0][2]["approved_shared_memberships"] == ()
    assert calls[1][0] == "native"
    assert calls[1][2]["preflight"] is preflight


def test_cli_stops_on_non_windows_without_running_preflight(monkeypatch, capsys):
    monkeypatch.setattr(cli.sys, "platform", "linux")
    monkeypatch.setattr(
        cli, "build_daily_input_projection_preflight",
        lambda *a, **k: (_ for _ in ()).throw(AssertionError("must not run")),
    )
    code = cli.main([
        "--source", "private.xlsm",
        "--service-date", "2026-11-23",
        "--baseline-date", "2026-01-05",
    ])
    assert code == 2
    assert "requires Windows" in capsys.readouterr().out


def test_cli_rejects_invalid_dates_without_opening_workbook(monkeypatch, capsys):
    monkeypatch.setattr(cli.sys, "platform", "win32")
    monkeypatch.setattr(
        cli, "build_daily_input_projection_preflight",
        lambda *a, **k: (_ for _ in ()).throw(AssertionError("must not run")),
    )
    code = cli.main([
        "--source", "private.xlsm",
        "--service-date", "not-a-date",
        "--baseline-date", "2026-01-05",
    ])
    assert code == 2
    assert "valid YYYY-MM-DD dates" in capsys.readouterr().out


def test_cli_hides_source_exception_details(monkeypatch, capsys):
    monkeypatch.setattr(cli.sys, "platform", "win32")
    monkeypatch.setattr(
        cli, "build_daily_input_projection_preflight",
        lambda *a, **k: (_ for _ in ()).throw(RuntimeError("FICTIONAL PATIENT PII")),
    )
    code = cli.main([
        "--source", "private.xlsm",
        "--service-date", "2026-11-23",
        "--baseline-date", "2026-01-05",
    ])
    output = capsys.readouterr().out
    assert code == 2
    assert "SAFETY STOP" in output
    assert "FICTIONAL PATIENT PII" not in output
