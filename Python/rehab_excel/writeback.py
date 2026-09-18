from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from hashlib import sha256
from pathlib import Path
from typing import Iterable

from .layout_contract import LayoutSafetyError, assert_cell_write_allowed, parse_cell
from .package_probe import WorkbookPackageProbe
from .source_contract import get_source_rule


class WritebackSafetyError(RuntimeError):
    """Raised when a requested Excel write violates the source contract."""


class WriteIntent(str, Enum):
    VALUE = "value"
    CLEAR = "clear"
    PRESENTATION = "presentation"


@dataclass(frozen=True)
class TextRunPatch:
    """Formatting for one 1-based character range inside a cell value."""

    start: int
    length: int
    strike_through: bool = False
    font_role: str | None = None

    def __post_init__(self) -> None:
        if self.start < 1:
            raise ValueError("TextRunPatch.start must be >= 1")
        if self.length < 1:
            raise ValueError("TextRunPatch.length must be >= 1")


@dataclass(frozen=True)
class CellPatch:
    """One future workbook mutation.

    The dry-run layer only describes and validates the intended change. Rich
    text formatting is represented as character runs so the native Excel
    writer can strike only the original line while keeping a replacement line
    active in the same cell when the UI requires it.
    """

    sheet: str
    cell: str
    value: object | None = None
    intent: WriteIntent = WriteIntent.VALUE
    strike_through: bool = False
    font_role: str | None = None
    fill_role: str | None = None
    border_role: str | None = None
    wrap_text: bool | None = None
    min_font_size: int | None = None
    text_runs: tuple[TextRunPatch, ...] = ()
    source_tag: str | None = None

    def normalized_key(self) -> tuple[str, str]:
        return self.sheet, self.cell.upper()


@dataclass(frozen=True)
class WritePlan:
    workbook_path: str
    patches: tuple[CellPatch, ...]

    @property
    def operation_count(self) -> int:
        return len(self.patches)

    @property
    def touched_sheets(self) -> tuple[str, ...]:
        return tuple(sorted({patch.sheet for patch in self.patches}))


@dataclass(frozen=True)
class DryRunReport:
    workbook_path: str
    operation_count: int
    touched_sheets: tuple[str, ...]
    sha256_before: str
    sha256_after: str
    workbook_unchanged: bool


def _file_sha256(path: str | Path) -> str:
    digest = sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _validate_cell_reference(cell: str) -> None:
    try:
        parse_cell(cell)
    except LayoutSafetyError as exc:
        raise WritebackSafetyError(str(exc)) from exc


def _validate_text_runs(patch: CellPatch) -> None:
    if not patch.text_runs:
        return
    if not isinstance(patch.value, str):
        raise WritebackSafetyError(
            f"Rich-text patch requires a string value: {patch.sheet}!{patch.cell}"
        )
    text_length = len(patch.value)
    previous_end = 0
    for run in patch.text_runs:
        end = run.start + run.length - 1
        if end > text_length:
            raise WritebackSafetyError(
                f"Text run exceeds cell value in {patch.sheet}!{patch.cell}: "
                f"{run.start}+{run.length - 1}>{text_length}"
            )
        if run.start <= previous_end:
            raise WritebackSafetyError(
                f"Overlapping or unsorted text runs in {patch.sheet}!{patch.cell}"
            )
        previous_end = end


def validate_write_plan(plan: WritePlan) -> None:
    """Validate a future write plan without modifying the workbook.

    Python may only target sheets explicitly marked as ``future_write_target``
    and cells inside verified write zones. Authoritative input sheets and
    legacy operational sheets remain protected.
    """

    workbook_path = Path(plan.workbook_path)
    if not workbook_path.exists():
        raise WritebackSafetyError(f"Workbook not found: {workbook_path}")

    workbook_sheets = set(WorkbookPackageProbe(workbook_path).sheet_names)

    seen: dict[tuple[str, str], CellPatch] = {}
    for patch in plan.patches:
        _validate_cell_reference(patch.cell)
        _validate_text_runs(patch)

        if patch.min_font_size is not None and patch.min_font_size < 1:
            raise WritebackSafetyError(
                f"Invalid minimum font size for {patch.sheet}!{patch.cell}"
            )

        if patch.sheet not in workbook_sheets:
            raise WritebackSafetyError(
                f"Write target sheet does not exist in workbook: {patch.sheet}"
            )

        rule = get_source_rule(patch.sheet)
        if rule is None or not rule.future_write_target:
            raise WritebackSafetyError(
                f"Sheet {patch.sheet} is protected from Python write-back"
            )

        try:
            assert_cell_write_allowed(patch.sheet, patch.cell)
        except LayoutSafetyError as exc:
            raise WritebackSafetyError(str(exc)) from exc

        key = patch.normalized_key()
        previous = seen.get(key)
        if previous is not None and previous != patch:
            raise WritebackSafetyError(
                f"Conflicting writes planned for {patch.sheet}!{patch.cell.upper()}"
            )
        seen[key] = patch


def build_write_plan(
    workbook_path: str | Path,
    patches: Iterable[CellPatch],
) -> WritePlan:
    plan = WritePlan(str(workbook_path), tuple(patches))
    validate_write_plan(plan)
    return plan


def dry_run_writeback(plan: WritePlan) -> DryRunReport:
    """Validate a write plan and prove the workbook bytes remain unchanged."""

    validate_write_plan(plan)
    before = _file_sha256(plan.workbook_path)

    # Deliberately no mutation. The plan is only inspected/validated.

    after = _file_sha256(plan.workbook_path)
    return DryRunReport(
        workbook_path=plan.workbook_path,
        operation_count=plan.operation_count,
        touched_sheets=plan.touched_sheets,
        sha256_before=before,
        sha256_after=after,
        workbook_unchanged=before == after,
    )
