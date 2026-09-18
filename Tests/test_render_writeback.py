from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook

from rehab_excel.render_plan import (
    CellRenderPlan,
    DailyExcelRenderPlan,
    RenderBorderRole,
    RenderFillRole,
    RenderFontRole,
    RenderIssue,
    RenderLine,
    RenderLineRole,
)
from rehab_excel.render_writeback import (
    RenderWritebackError,
    build_daily_write_plan,
    summarize_daily_write_plan,
)


def _workbook(tmp_path: Path) -> Path:
    path = tmp_path / "book.xlsx"
    wb = Workbook()
    wb.active.title = "THERAPIST_DAILY"
    wb.create_sheet("MASTER_SCHEDULE")
    wb.save(path)
    return path


def _render_plan(path: Path, cell: CellRenderPlan, issues=()):
    return DailyExcelRenderPlan(
        workbook_path=str(path),
        target_date=date(2026, 9, 18),
        bindings=(),
        cells=(cell,),
        issues=tuple(issues),
    )


def test_render_plan_becomes_rich_text_cell_patch(tmp_path):
    path = _workbook(tmp_path)
    cell = CellRenderPlan(
        sheet="THERAPIST_DAILY",
        cell="B2",
        lines=(
            RenderLine(
                "ΑΣΘΕΝΗΣ Α",
                RenderLineRole.ORIGINAL,
                strike_through=False,
                italic=True,
                font_role=RenderFontRole.MUTED,
            ),
            RenderLine(
                "ΑΣΘΕΝΗΣ Β",
                RenderLineRole.REPLACEMENT,
                strike_through=False,
                font_role=RenderFontRole.STUDENT_ACTIVE,
            ),
        ),
        fill_role=RenderFillRole.ROBOTIC,
        border_role=RenderBorderRole.INFECTIOUS,
        wrap_text=True,
        min_font_size=10,
    )
    plan = build_daily_write_plan(_render_plan(path, cell))
    patch = plan.patches[0]

    assert patch.value == "ΑΣΘΕΝΗΣ Α\nΑΣΘΕΝΗΣ Β"
    assert patch.fill_role == "robotic_pink"
    assert patch.border_role == "infectious_yellow"
    assert patch.wrap_text is True
    assert patch.min_font_size == 10
    assert len(patch.text_runs) == 2
    assert patch.text_runs[0].start == 1
    assert patch.text_runs[0].length == len("ΑΣΘΕΝΗΣ Α")
    assert patch.text_runs[0].strike_through is False
    assert patch.text_runs[0].italic is True
    assert patch.text_runs[1].start == len("ΑΣΘΕΝΗΣ Α") + 2
    assert patch.text_runs[1].font_role == "student_active_green"


def test_summary_counts_rich_and_struck_runs(tmp_path):
    path = _workbook(tmp_path)
    cell = CellRenderPlan(
        sheet="THERAPIST_DAILY",
        cell="C3",
        lines=(
            RenderLine("Α", RenderLineRole.ORIGINAL, strike_through=True),
            RenderLine("Β", RenderLineRole.REPLACEMENT),
        ),
    )
    plan = build_daily_write_plan(_render_plan(path, cell))
    summary = summarize_daily_write_plan(plan)
    assert summary.cell_count == 1
    assert summary.rich_text_cell_count == 1
    assert summary.struck_run_count == 1


def test_render_errors_block_write_plan(tmp_path):
    path = _workbook(tmp_path)
    cell = CellRenderPlan(
        sheet="THERAPIST_DAILY",
        cell="B2",
        lines=(RenderLine("Α", RenderLineRole.ACTIVE),),
    )
    render = _render_plan(
        path,
        cell,
        issues=(RenderIssue("BROKEN", "mapping failed", severity="error"),),
    )
    with pytest.raises(RenderWritebackError, match="BROKEN"):
        build_daily_write_plan(render)


def test_non_daily_sheet_is_not_emitted_yet(tmp_path):
    path = _workbook(tmp_path)
    cell = CellRenderPlan(
        sheet="MASTER_SCHEDULE",
        cell="D2",
        lines=(RenderLine("Α", RenderLineRole.ACTIVE),),
    )
    with pytest.raises(RenderWritebackError, match="Unsupported daily write target"):
        build_daily_write_plan(_render_plan(path, cell))
