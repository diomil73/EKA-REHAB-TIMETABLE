from pathlib import Path

from openpyxl import Workbook

from rehab_excel.daily_preview_composer import merge_presentation_patches
from rehab_excel.writeback import CellPatch, WriteIntent, build_write_plan


def _workbook(tmp_path: Path) -> Path:
    path = tmp_path / "preview.xlsm"
    xlsx = tmp_path / "preview.xlsx"
    wb = Workbook()
    wb.active.title = "THERAPIST_DAILY"
    wb.create_sheet("MASTER_SCHEDULE")
    wb.save(xlsx)
    path.write_bytes(xlsx.read_bytes())
    return path


def test_blue_presentation_merges_into_existing_value_patch(tmp_path):
    workbook = _workbook(tmp_path)
    base = build_write_plan(
        workbook,
        [
            CellPatch(
                "THERAPIST_DAILY",
                "B2",
                value="ΑΣΘΕΝΗΣ",
                source_tag="daily_text",
            )
        ],
    )
    blue = CellPatch(
        "THERAPIST_DAILY",
        "B2",
        intent=WriteIntent.PRESENTATION,
        fill_role="outpatient_light_blue",
        source_tag="outpatient_blue",
    )

    merged = merge_presentation_patches(base, (blue,))

    assert merged.operation_count == 1
    patch = merged.patches[0]
    assert patch.value == "ΑΣΘΕΝΗΣ"
    assert patch.intent == WriteIntent.VALUE
    assert patch.fill_role == "outpatient_light_blue"
    assert patch.source_tag == "daily_text+outpatient_blue"


def test_composer_rejects_non_presentation_overlay(tmp_path):
    workbook = _workbook(tmp_path)
    base = build_write_plan(
        workbook,
        [CellPatch("THERAPIST_DAILY", "B2", value="ΑΣΘΕΝΗΣ")],
    )
    invalid = CellPatch("THERAPIST_DAILY", "B2", value="ΑΛΛΟ")

    try:
        merge_presentation_patches(base, (invalid,))
    except Exception as exc:
        assert "PRESENTATION" in str(exc)
    else:
        raise AssertionError("Expected non-presentation patch to be rejected")
