from hashlib import sha256
from pathlib import Path

import pytest
from openpyxl import Workbook

from rehab_excel.writeback import (
    CellPatch,
    WritebackSafetyError,
    build_write_plan,
    dry_run_writeback,
)


def _hash(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def _workbook(tmp_path: Path) -> Path:
    path = tmp_path / "test.xlsx"
    wb = Workbook()
    default = wb.active
    default.title = "PATIENTS"
    for name in [
        "PATIENT_PLANNER",
        "SETTINGS",
        "THERAPIST_DAILY",
        "MASTER_SCHEDULE",
        "REPLACEMENTS",
        "CONFLICT_LOG",
        "REPLACEMENT_LOG",
        "STATISTICS",
    ]:
        wb.create_sheet(name)
    wb.save(path)
    return path


def test_dry_run_allows_confirmed_future_output_sheet(tmp_path):
    path = _workbook(tmp_path)
    plan = build_write_plan(
        path,
        [CellPatch("THERAPIST_DAILY", "B2", "replacement preview")],
    )
    assert plan.operation_count == 1
    assert plan.touched_sheets == ("THERAPIST_DAILY",)


def test_dry_run_protects_authoritative_patient_sheet(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(WritebackSafetyError, match="protected"):
        build_write_plan(path, [CellPatch("PATIENTS", "C2", "DO NOT WRITE")])


def test_dry_run_protects_base_schedule(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(WritebackSafetyError, match="protected"):
        build_write_plan(path, [CellPatch("PATIENT_PLANNER", "A2", "DO NOT WRITE")])


def test_dry_run_protects_legacy_operational_sheet(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(WritebackSafetyError, match="protected"):
        build_write_plan(path, [CellPatch("REPLACEMENT_LOG", "A2", "legacy")])


def test_dry_run_rejects_conflicting_same_cell_operations(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(WritebackSafetyError, match="Conflicting writes"):
        build_write_plan(
            path,
            [
                CellPatch("THERAPIST_DAILY", "B2", "one"),
                CellPatch("THERAPIST_DAILY", "b2", "two"),
            ],
        )


def test_dry_run_rejects_missing_sheet(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(WritebackSafetyError, match="does not exist"):
        build_write_plan(path, [CellPatch("NOT_A_SHEET", "A1", "x")])


def test_dry_run_does_not_change_workbook_bytes(tmp_path):
    path = _workbook(tmp_path)
    before = _hash(path)
    plan = build_write_plan(
        path,
        [
            CellPatch("THERAPIST_DAILY", "B2", "preview"),
            CellPatch("MASTER_SCHEDULE", "D2", "preview"),
        ],
    )
    report = dry_run_writeback(plan)
    after = _hash(path)

    assert report.workbook_unchanged is True
    assert report.sha256_before == report.sha256_after
    assert before == after
    assert report.operation_count == 2
    assert report.touched_sheets == ("MASTER_SCHEDULE", "THERAPIST_DAILY")


def test_dry_run_rejects_therapist_daily_header_cell(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(WritebackSafetyError, match="outside confirmed write zones"):
        build_write_plan(path, [CellPatch("THERAPIST_DAILY", "B1", "DO NOT OVERWRITE")])


def test_dry_run_rejects_master_schedule_patient_identity_cell(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(WritebackSafetyError, match="outside confirmed write zones"):
        build_write_plan(path, [CellPatch("MASTER_SCHEDULE", "C2", "DO NOT OVERWRITE")])


def test_dry_run_blocks_unmapped_statistics_cells(tmp_path):
    path = _workbook(tmp_path)
    # STATISTICS is a future target at sheet level, but no confirmed cell-level
    # write zone exists yet.
    with pytest.raises(WritebackSafetyError, match="no confirmed cell-level write zone"):
        build_write_plan(path, [CellPatch("STATISTICS", "A1", "blocked")])
