"""Build a fresh, private native-XLSM preview with an EMPTY afternoon sheet.

Never writes source, frozen evidence or the real clinical workbook. The new
sheet is technical/veryHidden; it is not a patient form or an approval to import.
Run only on Windows with Microsoft Excel and a fresh, disposable preview input.
"""
from __future__ import annotations

from hashlib import sha256
from pathlib import Path
import shutil
import sys

from rehab_excel.afternoon_slots_source import HEADERS, SHEET_NAME


class AfternoonPreviewError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_preview_paths(source: str | Path, destination: str | Path) -> tuple[Path, Path]:
    """Reject in-place writes, overwrite, non-XLSM, and symlinked targets."""
    src = Path(source).resolve(strict=True)
    dst_candidate = Path(destination).expanduser()
    if dst_candidate.is_symlink():
        raise AfternoonPreviewError("Output symlink forbidden")
    dst = dst_candidate.resolve()
    if (
        src == dst or src.suffix.lower() != ".xlsm"
        or dst.suffix.lower() != ".xlsm"
    ):
        raise AfternoonPreviewError("Distinct source/destination XLSM files required")
    if not src.is_file() or dst.exists():
        raise AfternoonPreviewError("Source missing or output already exists")
    if not dst.parent.is_dir():
        raise AfternoonPreviewError("Output directory does not exist")
    # Forensic evidence and original clinical workbooks are only source candidates
    # if explicitly chosen, but source bytes are never written by this tool.
    return src, dst


def build_afternoon_sheet_preview(source: str | Path, destination: str | Path) -> Path:
    """Install a header-only technical sheet on an unexisting output copy.

    Fail closed: if an afternoon sheet already exists, refuse instead of
    repurposing or clearing data. Verify the copy in a second Excel read-only
    session, and compare source SHA-256 before returning.
    """
    if sys.platform != "win32":
        raise AfternoonPreviewError("Native Microsoft Excel on Windows required")
    src, dst = validate_preview_paths(source, destination)
    original_digest = _sha256(src)
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise AfternoonPreviewError("pywin32 is required") from exc

    excel = None
    workbook = None
    created_output = False
    try:
        # Exclusive creation means a second process cannot silently overwrite.
        with dst.open("xb") as target, src.open("rb") as input_stream:
            created_output = True
            shutil.copyfileobj(input_stream, target)

        if _sha256(src) != original_digest:
            raise AfternoonPreviewError("SAFETY STOP: Source changed during copy")

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(
            str(dst), UpdateLinks=0, ReadOnly=False
        )
        if not bool(workbook.HasVBProject):
            raise AfternoonPreviewError("VBA project missing in preview")
        existing = {
            str(workbook.Worksheets(i).Name)
            for i in range(1, int(workbook.Worksheets.Count) + 1)
        }
        essentials = {"MASTER_SCHEDULE", "PATIENT_PLANNER", "PATIENTS"}
        if not essentials.issubset(existing):
            raise AfternoonPreviewError("Source copy is not a canonical Diomil app")
        if SHEET_NAME in existing:
            raise AfternoonPreviewError("Afternoon sheet already exists: refusing rewrite")

        master = workbook.Worksheets("MASTER_SCHEDULE")
        master_widths = tuple(
            float(master.Columns(i).ColumnWidth)
            for i in range(1, 13)
        )
        ws = workbook.Worksheets.Add(
            After=workbook.Worksheets(workbook.Worksheets.Count)
        )
        ws.Name = SHEET_NAME
        for index, header in enumerate(HEADERS, start=1):
            ws.Cells(1, index).Value = header
        ws.Visible = 2  # xlSheetVeryHidden
        if tuple(str(ws.Cells(1, i).Value) for i in range(1, 6)) != HEADERS:
            raise AfternoonPreviewError("Afternoon header readback mismatch")
        if any(ws.Cells(2, i).Value is not None for i in range(1, 6)):
            raise AfternoonPreviewError("Afternoon data is not empty")
        if tuple(float(master.Columns(i).ColumnWidth) for i in range(1, 13)) != master_widths:
            raise AfternoonPreviewError("MASTER widths changed before save")
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
        excel.Quit()
        excel = None

        check_excel = win32com.client.DispatchEx("Excel.Application")
        try:
            check_excel.Visible = False
            check_excel.DisplayAlerts = False
            check = check_excel.Workbooks.Open(
                str(dst), UpdateLinks=0, ReadOnly=True
            )
            try:
                ws = check.Worksheets(SHEET_NAME)
                if tuple(str(ws.Cells(1, i).Value) for i in range(1, 6)) != HEADERS:
                    raise AfternoonPreviewError("Reopened sheet schema mismatch")
                if any(ws.Cells(2, i).Value is not None for i in range(1, 6)):
                    raise AfternoonPreviewError("Reopened afternoon data not empty")
                if int(ws.Visible) != 2:
                    raise AfternoonPreviewError("Afternoon technical sheet not veryHidden")
                if not bool(check.HasVBProject):
                    raise AfternoonPreviewError("Reopened VBA project missing")
                master = check.Worksheets("MASTER_SCHEDULE")
                if tuple(float(master.Columns(i).ColumnWidth) for i in range(1, 13)) != master_widths:
                    raise AfternoonPreviewError("Reopened MASTER widths differ")
            finally:
                check.Close(SaveChanges=False)
        finally:
            check_excel.Quit()

        if _sha256(src) != original_digest:
            raise AfternoonPreviewError("SAFETY STOP: Source changed")
        return dst
    except Exception as exc:
        # Created output is disposable; never remove pre-existing files.
        # Retain failed outputs if Excel has already saved: they are forensic evidence.
        if created_output and workbook is None and excel is None:
            # The output might contain a saved VBA workbook; preserve it.
            pass
        if isinstance(exc, AfternoonPreviewError):
            raise
        raise AfternoonPreviewError("Native afternoon preview failed; preserve output for inspection") from exc
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
