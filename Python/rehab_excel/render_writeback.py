from __future__ import annotations

from dataclasses import dataclass

from .render_plan import DailyExcelRenderPlan
from .writeback import CellPatch, TextRunPatch, WriteIntent, WritePlan, build_write_plan


class RenderWritebackError(RuntimeError):
    """Raised when a semantic daily render plan cannot become a write plan."""


@dataclass(frozen=True)
class RenderWritebackSummary:
    cell_count: int
    rich_text_cell_count: int
    struck_run_count: int


def _line_runs(text_lines) -> tuple[TextRunPatch, ...]:
    runs: list[TextRunPatch] = []
    cursor = 1
    for index, line in enumerate(text_lines):
        if line.text:
            runs.append(
                TextRunPatch(
                    start=cursor,
                    length=len(line.text),
                    strike_through=line.strike_through,
                    italic=line.italic,
                    font_role=line.font_role.value,
                )
            )
        cursor += len(line.text)
        if index < len(text_lines) - 1:
            cursor += 1  # newline between render lines
    return tuple(runs)


def build_daily_write_plan(render_plan: DailyExcelRenderPlan) -> WritePlan:
    """Translate a verified DailyExcelRenderPlan into safe CellPatch objects.

    Only THERAPIST_DAILY cells are emitted at this stage. MASTER_SCHEDULE is
    intentionally left untouched so recurring day-pattern text remains the
    base source of truth.
    """

    if not render_plan.ok:
        details = "; ".join(
            f"{issue.code}: {issue.message}"
            for issue in render_plan.issues
            if issue.severity == "error"
        )
        raise RenderWritebackError(
            "Daily render plan contains errors and cannot be written: " + details
        )

    patches = []
    for cell in render_plan.cells:
        if cell.sheet != "THERAPIST_DAILY":
            raise RenderWritebackError(
                f"Unsupported daily write target at this stage: {cell.sheet}!{cell.cell}"
            )
        patches.append(
            CellPatch(
                sheet=cell.sheet,
                cell=cell.cell,
                value=cell.text,
                intent=WriteIntent.PRESENTATION,
                fill_role=cell.fill_role.value,
                border_role=cell.border_role.value,
                wrap_text=cell.wrap_text,
                min_font_size=cell.min_font_size,
                text_runs=_line_runs(cell.lines),
                source_tag=f"daily_state:{render_plan.target_date.isoformat()}",
            )
        )

    return build_write_plan(render_plan.workbook_path, patches)


def summarize_daily_write_plan(plan: WritePlan) -> RenderWritebackSummary:
    return RenderWritebackSummary(
        cell_count=len(plan.patches),
        rich_text_cell_count=sum(bool(patch.text_runs) for patch in plan.patches),
        struck_run_count=sum(
            1
            for patch in plan.patches
            for run in patch.text_runs
            if run.strike_through
        ),
    )
