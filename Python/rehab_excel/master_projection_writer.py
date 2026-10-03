from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Protocol

from openpyxl import load_workbook

from .master_projection import (
    MasterProjectionRow,
    build_master_projection,
    remap_master_formula,
)
from .native_excel import NativeStylePalette


class MasterProjectionWriteError(RuntimeError):
    """Raised when a safe MASTER_SCHEDULE projection preview cannot be produced."""


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


class MasterProjectionBackend(Protocol):
    def apply(
        self,
        workbook_path: Path,
        projection: tuple[MasterProjectionRow, ...],
    ) -> None:
        ...


class Win32ComMasterProjectionBackend:
    """Rebuild MASTER_SCHEDULE as a sorted projection of PATIENT_PLANNER."""

    def __init__(self, palette: NativeStylePalette | None = None) -> None:
        self.palette = palette or NativeStylePalette()

    @staticmethod
    def _last_value_row(ws) -> int:
        rows = []
        for col in (1, 2, 3, 11):
            try:
                rows.append(int(ws.Cells(ws.Rows.Count, col).End(-4162).Row))
            except Exception:
                pass
        return max(rows or [1])

    @staticmethod
    def _template_formulas(ws) -> dict[int, str]:
        formulas: dict[int, str] = {}
        for col in range(1, 12):
            value = ws.Cells(2, col).Formula
            text = str(value or "")
            if text.startswith("="):
                formulas[col] = text
        required = set(range(1, 12))
        missing = sorted(required.difference(formulas))
        if missing:
            raise MasterProjectionWriteError(
                f"MASTER_SCHEDULE row 2 is missing formula templates in columns: {missing}"
            )
        return formulas

    def _apply_infectious_style(self, row_range) -> None:
        row_range.Interior.Color = self.palette.infectious_yellow
        for edge in (7, 8, 9, 10):
            border = row_range.Borders(edge)
            border.LineStyle = 1
            border.Weight = 2
            border.Color = self.palette.infectious_yellow

    def apply(
        self,
        workbook_path: Path,
        projection: tuple[MasterProjectionRow, ...],
    ) -> None:
        if sys.platform != "win32":
            raise MasterProjectionWriteError(
                "MASTER projection write-back requires Windows with Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise MasterProjectionWriteError(
                "pywin32 is required for MASTER projection write-back"
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
            try:
                ws = workbook.Worksheets("MASTER_SCHEDULE")
            except Exception as exc:
                raise MasterProjectionWriteError(
                    "Workbook has no MASTER_SCHEDULE sheet"
                ) from exc

            templates = self._template_formulas(ws)
            old_last = self._last_value_row(ws)
            new_last = max(1, len(projection) + 1)
            clear_last = max(old_last, new_last)

            # Reset old presentation/content using row 2 as the neutral template.
            for row in range(2, clear_last + 1):
                try:
                    ws.Range("A2:K2").Copy()
                    ws.Range(f"A{row}:K{row}").PasteSpecial(Paste=-4122)  # xlPasteFormats
                except Exception:
                    pass
                ws.Range(f"A{row}:K{row}").ClearContents()

            for item in projection:
                row = item.target_row
                for col in range(1, 12):
                    ws.Cells(row, col).Formula = remap_master_formula(
                        templates[col],
                        item.planner_row,
                    )

                if item.infectious:
                    self._apply_infectious_style(ws.Range(f"B{row}:C{row}"))

            workbook.Application.CutCopyMode = False
            excel.CalculateFull()
            workbook.Save()
        except MasterProjectionWriteError:
            raise
        except Exception as exc:
            raise MasterProjectionWriteError(
                f"Excel MASTER projection failed: {exc}"
            ) from exc
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


@dataclass(frozen=True)
class MasterProjectionPreviewReport:
    source_path: str
    output_path: str
    projected_rows: int
    source_unchanged: bool
    verified_in_output: bool


def _verify_projection(
    output: Path,
    projection: tuple[MasterProjectionRow, ...],
) -> bool:
    wb = load_workbook(
        output,
        read_only=False,
        data_only=False,
        keep_vba=True,
    )
    try:
        if "MASTER_SCHEDULE" not in wb.sheetnames:
            return False
        ws = wb["MASTER_SCHEDULE"]
        for item in projection:
            row = item.target_row
            expected = {
                1: f"=PATIENT_PLANNER!D{item.planner_row}",
                2: f"=PATIENT_PLANNER!B{item.planner_row}",
                3: f"=PATIENT_PLANNER!C{item.planner_row}",
                11: f"=PATIENT_PLANNER!E{item.planner_row}",
            }
            for col, formula in expected.items():
                if str(ws.cell(row, col).value or "") != formula:
                    return False

        tail_row = len(projection) + 2
        if tail_row <= ws.max_row:
            for col in range(1, 12):
                if ws.cell(tail_row, col).value not in (None, ""):
                    return False
        return True
    finally:
        wb.close()


def create_master_projection_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    backend: MasterProjectionBackend | None = None,
    overwrite: bool = False,
) -> MasterProjectionPreviewReport:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise MasterProjectionWriteError(f"Source workbook not found: {source}")
    if source == output:
        raise MasterProjectionWriteError("Output must differ from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise MasterProjectionWriteError("Source and output must both be .xlsm")
    if output.exists() and not overwrite:
        raise MasterProjectionWriteError(f"Output already exists: {output}")

    projection = build_master_projection(source)
    source_before = _sha256(source)

    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        try:
            output.unlink()
        except PermissionError as exc:
            raise MasterProjectionWriteError(
                f"Preview workbook is open or locked: {output}"
            ) from exc
    shutil.copy2(source, output)

    selected = backend or Win32ComMasterProjectionBackend()
    try:
        selected.apply(output, projection)
    except Exception:
        try:
            output.unlink()
        except OSError:
            pass
        raise

    source_after = _sha256(source)
    if source_before != source_after:
        try:
            output.unlink()
        except OSError:
            pass
        raise MasterProjectionWriteError(
            "Source workbook changed during MASTER projection preview"
        )

    verified = _verify_projection(output, projection)
    if not verified:
        try:
            output.unlink()
        except OSError:
            pass
        raise MasterProjectionWriteError(
            "MASTER projection verification failed after write-back"
        )

    return MasterProjectionPreviewReport(
        source_path=str(source),
        output_path=str(output),
        projected_rows=len(projection),
        source_unchanged=True,
        verified_in_output=True,
    )
