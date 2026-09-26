from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Protocol

from .writeback import CellPatch, WriteIntent, WritePlan, validate_write_plan


class NativeExcelWriteError(RuntimeError):
    """Raised when a native Excel copy-write cannot be completed safely."""


def _sha256(path: str | Path) -> str:
    digest = sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def excel_rgb(red: int, green: int, blue: int) -> int:
    """Return the OLE color integer produced by VBA's RGB(r, g, b)."""

    for component in (red, green, blue):
        if not 0 <= component <= 255:
            raise ValueError("RGB components must be between 0 and 255")
    return red | (green << 8) | (blue << 16)


@dataclass(frozen=True)
class NativeStylePalette:
    """Semantic colors used by the current v27 presentation contract.

    Infectious yellow and robotic salmon/pink are taken from existing workbook
    styles rather than invented during Python migration. Student green and
    outpatient blue are dedicated future presentation roles.
    """

    infectious_yellow: int = excel_rgb(255, 255, 67)   # FFFFFF43
    robotic_pink: int = excel_rgb(248, 203, 173)       # FFF8CBAD in v27
    outpatient_light_blue: int = excel_rgb(221, 235, 247)  # DDEBF7
    robotic_orange: int = excel_rgb(255, 140, 0)       # FFFF8C00 in v27 font
    student_green: int = excel_rgb(0, 128, 0)
    muted_gray: int = excel_rgb(102, 102, 102)
    default_black: int = excel_rgb(0, 0, 0)


class ExcelPatchBackend(Protocol):
    """Backend contract so native Excel behavior can be unit-tested safely."""

    def apply(self, workbook_path: Path, patches: tuple[CellPatch, ...]) -> None:
        ...


@dataclass(frozen=True)
class NativeWriteReport:
    source_path: str
    output_path: str
    operation_count: int
    source_sha256_before: str
    source_sha256_after: str
    output_sha256: str
    source_unchanged: bool


class Win32ComExcelBackend:
    """Apply CellPatch objects using Microsoft Excel itself on Windows.

    The workbook is opened only after a copy has been created. Excel performs
    the save, preserving VBA, drawings and native workbook structures that are
    unsafe to rewrite with openpyxl.
    """

    def __init__(self, palette: NativeStylePalette | None = None) -> None:
        self.palette = palette or NativeStylePalette()

    @staticmethod
    def _font_color_for_role(role: str | None, palette: NativeStylePalette) -> int | None:
        return {
            "student_active_green": palette.student_green,
            "robotic_orange": palette.robotic_orange,
            "muted": palette.muted_gray,
            "default": palette.default_black,
        }.get(role)

    @staticmethod
    def _characters(cell, start: int, length: int):
        # With pywin32's dynamic Excel dispatch, Range.Characters(Start, Length)
        # is exposed as an indexed COM property and calling it directly can raise
        # DISP_E_MEMBERNOTFOUND. pywin32 exposes the property getter as
        # GetCharacters(Start, Length), which is the compatible route on these
        # Office installs. Keep a direct Characters fallback for generated/static
        # wrappers where that syntax is callable.
        getter = getattr(cell, "GetCharacters", None)
        if callable(getter):
            try:
                return getter(start, length)
            except Exception as get_exc:
                direct = getattr(cell, "Characters", None)
                if callable(direct):
                    try:
                        return direct(start, length)
                    except Exception:
                        pass
                raise NativeExcelWriteError(
                    f"Excel rich-text character access failed "
                    f"(start={start}, length={length}): {get_exc}"
                ) from get_exc

        direct = getattr(cell, "Characters", None)
        if callable(direct):
            try:
                return direct(start, length)
            except Exception as exc:
                raise NativeExcelWriteError(
                    f"Excel rich-text character access failed "
                    f"(start={start}, length={length}): {exc}"
                ) from exc

        raise NativeExcelWriteError(
            "Excel Range exposes neither GetCharacters nor callable Characters"
        )

    def _apply_patch(self, workbook, patch: CellPatch) -> None:
        worksheet = workbook.Worksheets(patch.sheet)
        cell = worksheet.Range(patch.cell)

        if patch.intent == WriteIntent.CLEAR:
            cell.ClearContents()
        else:
            cell.Value = patch.value

        if patch.wrap_text is not None:
            cell.WrapText = bool(patch.wrap_text)

        if patch.min_font_size is not None:
            try:
                current_size = float(cell.Font.Size)
            except (TypeError, ValueError):
                current_size = 0
            if current_size < patch.min_font_size:
                cell.Font.Size = patch.min_font_size

        if patch.strike_through:
            cell.Font.Strikethrough = True

        whole_color = self._font_color_for_role(patch.font_role, self.palette)
        if whole_color is not None:
            cell.Font.Color = whole_color

        if patch.fill_role == "infectious_yellow":
            cell.Interior.Color = self.palette.infectious_yellow
        elif patch.fill_role == "robotic_pink":
            cell.Interior.Color = self.palette.robotic_pink
        elif patch.fill_role == "outpatient_light_blue":
            cell.Interior.Color = self.palette.outpatient_light_blue

        if patch.border_role == "infectious_yellow":
            # xlContinuous = 1, xlThin = 2. Apply the four outside edges only.
            for edge in (7, 8, 9, 10):  # left, top, bottom, right
                border = cell.Borders(edge)
                border.LineStyle = 1
                border.Weight = 2
                border.Color = self.palette.infectious_yellow

        # Reset whole-string rich-text state first. Then format individual
        # lines/runs. This prevents stale presentation from a previous day.
        if patch.text_runs:
            cell.Font.Strikethrough = False
            cell.Font.Italic = False
            for run in patch.text_runs:
                chars = self._characters(cell, run.start, run.length)
                chars.Font.Strikethrough = bool(run.strike_through)
                chars.Font.Italic = bool(run.italic)
                run_color = self._font_color_for_role(run.font_role, self.palette)
                if run_color is not None:
                    chars.Font.Color = run_color

    def apply(self, workbook_path: Path, patches: tuple[CellPatch, ...]) -> None:
        if sys.platform != "win32":
            raise NativeExcelWriteError(
                "Native Excel write-back requires Windows with Microsoft Excel installed"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise NativeExcelWriteError(
                "pywin32 is required for native Excel write-back. "
                "Install with: pip install pywin32"
            ) from exc

        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            workbook = excel.Workbooks.Open(
                str(workbook_path.resolve()),
                UpdateLinks=0,
                ReadOnly=False,
            )
            for patch in patches:
                self._apply_patch(workbook, patch)
            workbook.Save()
        except Exception as exc:
            raise NativeExcelWriteError(f"Excel write-back failed: {exc}") from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass


def apply_write_plan_to_copy(
    plan: WritePlan,
    output_path: str | Path,
    *,
    backend: ExcelPatchBackend | None = None,
    overwrite: bool = False,
) -> NativeWriteReport:
    """Apply a validated plan to a NEW workbook copy, never to the source.

    The source workbook is hashed before and after the operation. If the source
    bytes change for any reason, the function raises even if Excel saved the
    output successfully.
    """

    validate_write_plan(plan)
    source = Path(plan.workbook_path).resolve()
    output = Path(output_path).resolve()

    if source == output:
        raise NativeExcelWriteError("Output path must be different from source workbook")
    if source.suffix.lower() != ".xlsm" or output.suffix.lower() != ".xlsm":
        raise NativeExcelWriteError("Native copy-write requires .xlsm source and output")
    if output.exists() and not overwrite:
        raise NativeExcelWriteError(f"Output workbook already exists: {output}")

    output.parent.mkdir(parents=True, exist_ok=True)
    source_before = _sha256(source)

    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    selected_backend = backend or Win32ComExcelBackend()
    try:
        selected_backend.apply(output, plan.patches)
    except Exception:
        # Never leave a file that looks like a successful preview when native
        # Excel failed part-way through the operation.
        output.unlink(missing_ok=True)
        raise

    source_after = _sha256(source)
    if source_before != source_after:
        output.unlink(missing_ok=True)
        raise NativeExcelWriteError("Source workbook changed during copy-write operation")
    if not output.exists():
        raise NativeExcelWriteError("Excel backend did not produce an output workbook")

    return NativeWriteReport(
        source_path=str(source),
        output_path=str(output),
        operation_count=plan.operation_count,
        source_sha256_before=source_before,
        source_sha256_after=source_after,
        output_sha256=_sha256(output),
        source_unchanged=True,
    )
