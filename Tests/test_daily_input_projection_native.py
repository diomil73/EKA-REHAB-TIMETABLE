"""Synthetic composition tests; native COM diagnostics are monkeypatched."""
from datetime import date
from hashlib import sha256

import pytest
from openpyxl import Workbook

import rehab_excel.daily_input_projection_native as composite
from rehab_excel.daily_input_input_ledger_native import InputLedgerNativeReadback
from rehab_excel.daily_input_master_projection_native import MasterProjectionNativeReadback
from rehab_excel.master_projection import MasterProjectionRow
from rehab_excel.daily_input_therapist_projection_expected import TherapistDailyExpectedGrid
from rehab_excel.daily_input_therapist_projection_native import (
    TherapistDailyCellExpectation,
    TherapistDailyNativeReadback,
)
from rehab_excel.daily_input_projection_native import (
    DailyInputProjectionNativeReadbackError,
    verify_daily_input_projection_native_readonly,
)


SERVICE_DATE = date(2026, 11, 23)


def _source(tmp_path):
    path = tmp_path / "composite-projection.xlsm"
    wb = Workbook()
    wb.active.title = "MASTER_SCHEDULE"
    wb.save(path)
    wb.close()
    return path


def _evidence(path):
    digest = sha256(path.read_bytes()).hexdigest()
    headers = tuple(f"THERAPIST {i}" for i in range(1, 10)) * 2
    cells = tuple(
        TherapistDailyCellExpectation(
            f"{chr(ord('B') + col)}{row}", None,
        )
        for row in (*range(2, 9), *range(12, 19))
        for col in range(9)
    )
    grid = TherapistDailyExpectedGrid(
        service_date=SERVICE_DATE,
        source_workbook_sha256=digest,
        provider_headers=headers,
        cells=cells,
        expected_assignments=(),
        occupied_cells=0,
    )
    ledger = InputLedgerNativeReadback(
        service_date=SERVICE_DATE,
        workbook_sha256=digest,
        source_change_count=2,
        active_ledger_count=2,
        native_excel_version="16.0",
    )
    master = MasterProjectionNativeReadback(
        workbook_sha256=digest,
        projected_rows=0,
        header_row=3,
        status_column=11,
        native_excel_version="16.0",
    )
    therapist = TherapistDailyNativeReadback(
        workbook_sha256=digest,
        checked_provider_headers=18,
        checked_grid_cells=126,
        occupied_grid_cells=0,
        native_excel_version="16.0",
    )
    return digest, grid, ledger, master, therapist


def _patch_readers(monkeypatch, master, therapist):
    calls = []

    def master_reader(path, **kwargs):
        calls.append(("master", path, kwargs))
        return master

    def therapist_reader(path, **kwargs):
        calls.append(("therapist", path, kwargs))
        return therapist

    monkeypatch.setattr(composite, "verify_master_projection_native_readonly", master_reader)
    monkeypatch.setattr(composite, "verify_therapist_daily_native_readonly", therapist_reader)
    return calls


def test_composite_binds_input_ledger_master_and_full_therapist_grid(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, master, therapist = _evidence(path)
    calls = _patch_readers(monkeypatch, master, therapist)

    result = verify_daily_input_projection_native_readonly(
        path,
        service_date=SERVICE_DATE,
        input_ledger_readback=ledger,
        expected_projection=(), expected_projection_source_sha256=digest,
        expected_therapist_grid=grid,
    )

    assert result.workbook_sha256 == digest
    assert result.service_date == SERVICE_DATE
    assert result.master_projected_rows == 0
    assert result.therapist_grid_cells_checked == 126
    assert result.therapist_occupied_cells == 0
    assert result.input_ledger_source_changes == 2
    assert result.input_ledger_active_rows == 2
    assert len(calls) == 2
    assert calls[0][0] == "master"
    assert calls[1][0] == "therapist"
    assert calls[1][2]["expected_cells"] == grid.cells


@pytest.mark.parametrize("mismatch", ["ledger-hash", "grid-hash", "ledger-date", "grid-date"])
def test_composite_rejects_cross_source_or_date_binding_mismatch(tmp_path, monkeypatch, mismatch):
    path = _source(tmp_path)
    digest, grid, ledger, master, therapist = _evidence(path)
    calls = _patch_readers(monkeypatch, master, therapist)
    if mismatch == "ledger-hash":
        ledger = InputLedgerNativeReadback(SERVICE_DATE, "0" * 64, 2, 2, "16.0")
    elif mismatch == "grid-hash":
        grid = TherapistDailyExpectedGrid(
            SERVICE_DATE, "0" * 64, grid.provider_headers, grid.cells,
            grid.expected_assignments, grid.occupied_cells,
        )
    elif mismatch == "ledger-date":
        ledger = InputLedgerNativeReadback(date(2026, 11, 24), digest, 2, 2, "16.0")
    else:
        grid = TherapistDailyExpectedGrid(
            date(2026, 11, 24), digest, grid.provider_headers, grid.cells,
            grid.expected_assignments, grid.occupied_cells,
        )

    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )
    assert calls == []


def test_composite_rejects_master_or_therapist_digest_drift(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, master, therapist = _evidence(path)
    wrong_master = MasterProjectionNativeReadback("0" * 64, 4, 3, 11, "16.0")
    _patch_readers(monkeypatch, wrong_master, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )

    wrong_therapist = TherapistDailyNativeReadback("0" * 64, 18, 126, 0, "16.0")
    _patch_readers(monkeypatch, master, wrong_therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )


def test_composite_rejects_workbook_mutation_between_readbacks(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, master, therapist = _evidence(path)
    _patch_readers(monkeypatch, master, therapist)

    def mutate_after_readback(path, **kwargs):
        path.write_bytes(path.read_bytes() + b"tamper")
        return therapist

    monkeypatch.setattr(composite, "verify_therapist_daily_native_readonly", mutate_after_readback)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )


def test_composite_rejects_inconsistent_component_counts(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, master, therapist = _evidence(path)
    wrong_master = MasterProjectionNativeReadback(digest, 1, 3, 11, "16.0")
    _patch_readers(monkeypatch, wrong_master, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )

    wrong_therapist = TherapistDailyNativeReadback(digest, 18, 126, 1, "16.0")
    _patch_readers(monkeypatch, master, wrong_therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )

    wrong_ledger = InputLedgerNativeReadback(SERVICE_DATE, digest, 2, 1, "16.0")
    _patch_readers(monkeypatch, master, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=wrong_ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )


def test_composite_rejects_boolean_or_missing_result_metadata(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, master, therapist = _evidence(path)
    bad_master = MasterProjectionNativeReadback(digest, True, 3, 11, "16.0")
    _patch_readers(monkeypatch, bad_master, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )

    bad_therapist = TherapistDailyNativeReadback(digest, 18, 126, True, "16.0")
    _patch_readers(monkeypatch, master, bad_therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )

    bad_ledger = InputLedgerNativeReadback(SERVICE_DATE, digest, 2, 2, None)
    _patch_readers(monkeypatch, master, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=bad_ledger,
            expected_projection=(), expected_projection_source_sha256=digest, expected_therapist_grid=grid,
        )


def test_composite_requires_master_preflight_digest_to_match_source(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, master, therapist = _evidence(path)
    calls = _patch_readers(monkeypatch, master, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256="0" * 64,
            expected_therapist_grid=grid,
        )
    assert calls == []

    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=(), expected_projection_source_sha256="not-a-sha",
            expected_therapist_grid=grid,
        )
    assert calls == []


def test_composite_requires_complete_master_formula_and_evaluated_value_coverage(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, _master, therapist = _evidence(path)
    expected = (
        MasterProjectionRow(5, 2, "P-001", "A01", "FICTIONAL PATIENT",
                            False, "active", None),
    )
    master = MasterProjectionNativeReadback(
        digest, 1, 3, 11, "16.0",
        formula_cells_checked=11, evaluated_cells_checked=11,
    )
    _patch_readers(monkeypatch, master, therapist)
    result = verify_daily_input_projection_native_readonly(
        path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
        expected_projection=expected, expected_projection_source_sha256=digest,
        expected_therapist_grid=grid,
    )
    assert result.master_formula_cells_checked == 11
    assert result.master_evaluated_cells_checked == 11

    short_master = MasterProjectionNativeReadback(
        digest, 1, 3, 11, "16.0",
        formula_cells_checked=10, evaluated_cells_checked=11,
    )
    _patch_readers(monkeypatch, short_master, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=expected, expected_projection_source_sha256=digest,
            expected_therapist_grid=grid,
        )


def test_composite_accounts_for_literal_doctor_identity_in_master_coverage(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, _master, therapist = _evidence(path)
    expected = (
        MasterProjectionRow(5, 2, "P-001", "A01", "FICTIONAL PATIENT",
                            False, "active", "DOCTOR X"),
    )
    master = MasterProjectionNativeReadback(
        digest, 1, 3, 11, "16.0",
        formula_cells_checked=10, evaluated_cells_checked=11,
    )
    _patch_readers(monkeypatch, master, therapist)
    result = verify_daily_input_projection_native_readonly(
        path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
        expected_projection=expected, expected_projection_source_sha256=digest,
        expected_therapist_grid=grid,
    )
    assert result.master_formula_cells_checked == 10
    assert result.master_evaluated_cells_checked == 11

def test_composite_requires_independent_psychology_literal_in_expanded_master(tmp_path, monkeypatch):
    path = _source(tmp_path)
    digest, grid, ledger, _master, therapist = _evidence(path)
    expected = (
        MasterProjectionRow(
            5, 2, "P-001", "A01", "FICTIONAL PATIENT", False, "active",
            psychology_display="16:30\\nΔΕ, ΠΕ\\nSYNTHETIC PROVIDER",
        ),
    )
    master = MasterProjectionNativeReadback(
        digest, 1, 3, 13, "16.0",
        formula_cells_checked=11, evaluated_cells_checked=12,
    )
    _patch_readers(monkeypatch, master, therapist)
    result = verify_daily_input_projection_native_readonly(
        path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
        expected_projection=expected, expected_projection_source_sha256=digest,
        expected_therapist_grid=grid,
    )
    assert result.master_formula_cells_checked == 11
    assert result.master_evaluated_cells_checked == 12

    incomplete = MasterProjectionNativeReadback(
        digest, 1, 3, 13, "16.0",
        formula_cells_checked=11, evaluated_cells_checked=11,
    )
    _patch_readers(monkeypatch, incomplete, therapist)
    with pytest.raises(DailyInputProjectionNativeReadbackError):
        verify_daily_input_projection_native_readonly(
            path, service_date=SERVICE_DATE, input_ledger_readback=ledger,
            expected_projection=expected, expected_projection_source_sha256=digest,
            expected_therapist_grid=grid,
        )
