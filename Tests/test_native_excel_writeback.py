from hashlib import sha256
from pathlib import Path

import pytest
from openpyxl import Workbook

from rehab_excel.native_excel import (
    NativeExcelWriteError,
    Win32ComExcelBackend,
    apply_write_plan_to_copy,
    excel_rgb,
)
from rehab_excel.writeback import CellPatch, TextRunPatch, WritebackSafetyError, build_write_plan


def _hash(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def _xlsm_container(tmp_path: Path) -> Path:
    # The safety layer probes workbook XML and does not require VBA for this unit
    # test. The extension mirrors the production copy-write contract.
    xlsx = tmp_path / "seed.xlsx"
    wb = Workbook()
    wb.active.title = "THERAPIST_DAILY"
    wb.create_sheet("MASTER_SCHEDULE")
    wb.save(xlsx)
    target = tmp_path / "source.xlsm"
    target.write_bytes(xlsx.read_bytes())
    return target


class FakeBackend:
    def __init__(self, marker=b"\nFAKE-EXCEL-SAVE"):
        self.marker = marker
        self.calls = []

    def apply(self, workbook_path: Path, patches):
        self.calls.append((workbook_path, patches))
        workbook_path.write_bytes(workbook_path.read_bytes() + self.marker)


class FailingBackend:
    def apply(self, workbook_path: Path, patches):
        workbook_path.write_bytes(workbook_path.read_bytes() + b"partial")
        raise RuntimeError("simulated Excel failure")


def test_excel_rgb_matches_vba_layout():
    assert excel_rgb(255, 255, 67) == 67 * 65536 + 255 * 256 + 255


def test_copy_write_never_changes_source(tmp_path):
    source = _xlsm_container(tmp_path)
    before = _hash(source)
    plan = build_write_plan(source, [CellPatch("THERAPIST_DAILY", "B2", "preview")])
    output = tmp_path / "preview.xlsm"
    backend = FakeBackend()

    report = apply_write_plan_to_copy(plan, output, backend=backend)

    assert _hash(source) == before
    assert report.source_unchanged is True
    assert report.source_sha256_before == report.source_sha256_after
    assert output.exists()
    assert _hash(output) != before
    assert backend.calls[0][0] == output.resolve()


def test_copy_write_rejects_source_as_destination(tmp_path):
    source = _xlsm_container(tmp_path)
    plan = build_write_plan(source, [CellPatch("THERAPIST_DAILY", "B2", "x")])
    with pytest.raises(NativeExcelWriteError, match="different from source"):
        apply_write_plan_to_copy(plan, source, backend=FakeBackend())


def test_copy_write_refuses_existing_output_by_default(tmp_path):
    source = _xlsm_container(tmp_path)
    output = tmp_path / "preview.xlsm"
    output.write_bytes(b"existing")
    plan = build_write_plan(source, [CellPatch("THERAPIST_DAILY", "B2", "x")])
    with pytest.raises(NativeExcelWriteError, match="already exists"):
        apply_write_plan_to_copy(plan, output, backend=FakeBackend())
    assert output.read_bytes() == b"existing"


def test_failed_native_write_removes_partial_preview(tmp_path):
    source = _xlsm_container(tmp_path)
    output = tmp_path / "preview.xlsm"
    plan = build_write_plan(source, [CellPatch("THERAPIST_DAILY", "B2", "x")])
    with pytest.raises(RuntimeError, match="simulated Excel failure"):
        apply_write_plan_to_copy(plan, output, backend=FailingBackend())
    assert not output.exists()


def test_rich_text_runs_are_validated(tmp_path):
    source = _xlsm_container(tmp_path)
    with pytest.raises(WritebackSafetyError, match="exceeds cell value"):
        build_write_plan(
            source,
            [
                CellPatch(
                    "THERAPIST_DAILY",
                    "B2",
                    "abc",
                    text_runs=(TextRunPatch(1, 4),),
                )
            ],
        )


def test_overlapping_rich_text_runs_are_rejected(tmp_path):
    source = _xlsm_container(tmp_path)
    with pytest.raises(WritebackSafetyError, match="Overlapping"):
        build_write_plan(
            source,
            [
                CellPatch(
                    "THERAPIST_DAILY",
                    "B2",
                    "abcdef",
                    text_runs=(TextRunPatch(1, 3), TextRunPatch(3, 2)),
                )
            ],
        )


class _DynamicCharactersCell:
    def __init__(self):
        self.get_calls = []

    def GetCharacters(self, *args):
        self.get_calls.append(args)
        return "chars"


def test_characters_prefers_pywin32_getcharacters():
    cell = _DynamicCharactersCell()
    result = Win32ComExcelBackend._characters(cell, 2, 5)
    assert result == "chars"
    assert cell.get_calls == [(2, 5)]


class _StaticCharactersCell:
    def __init__(self):
        self.calls = []

    def Characters(self, *args):
        self.calls.append(args)
        return "static-chars"


def test_characters_falls_back_to_callable_characters():
    cell = _StaticCharactersCell()
    result = Win32ComExcelBackend._characters(cell, 3, 4)
    assert result == "static-chars"
    assert cell.calls == [(3, 4)]
