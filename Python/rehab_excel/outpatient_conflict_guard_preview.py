from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys

from .outpatient_conflict_guard_vba import (
    FORM_NAME,
    install_outpatient_conflict_guard,
)
from .scheduling_conflict_popup_vba import (
    SCHEDULING_CONFLICT_FORM_NAME,
    SCHEDULING_CONFLICT_MODULE_NAME,
)


class OutpatientConflictGuardPreviewError(RuntimeError):
    pass


@dataclass(frozen=True)
class OutpatientConflictGuardPreviewReport:
    source_path: str
    output_path: str
    source_unchanged: bool
    outpatient_form_present: bool
    popup_form_present: bool
    popup_module_present: bool


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _safe_unlink(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except PermissionError:
        pass


def create_outpatient_conflict_guard_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> OutpatientConflictGuardPreviewReport:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise OutpatientConflictGuardPreviewError(f"Source workbook not found: {source}")
    if source == output:
        raise OutpatientConflictGuardPreviewError("Output must differ from source")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise OutpatientConflictGuardPreviewError("Source and output must both be .xlsm")
    if output.exists() and not overwrite:
        raise OutpatientConflictGuardPreviewError(f"Output already exists: {output}")
    if sys.platform != "win32":
        raise OutpatientConflictGuardPreviewError("Windows Excel is required")

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        _safe_unlink(output)
        raise OutpatientConflictGuardPreviewError("pywin32 is required") from exc

    excel = None
    workbook = None
    failure: Exception | None = None
    outpatient_form_present = False
    popup_form_present = False
    popup_module_present = False

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject
        install_outpatient_conflict_guard(vbproject)
        workbook.Save()

        names = {
            vbproject.VBComponents(index).Name
            for index in range(1, vbproject.VBComponents.Count + 1)
        }
        outpatient_form_present = FORM_NAME in names
        popup_form_present = SCHEDULING_CONFLICT_FORM_NAME in names
        popup_module_present = SCHEDULING_CONFLICT_MODULE_NAME in names
        if not (outpatient_form_present and popup_form_present and popup_module_present):
            raise OutpatientConflictGuardPreviewError(
                "Conflict guard components were not all installed"
            )
    except Exception as exc:
        failure = exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
            workbook = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None

    if failure is not None:
        _safe_unlink(output)
        if isinstance(failure, OutpatientConflictGuardPreviewError):
            raise failure
        raise OutpatientConflictGuardPreviewError(
            f"Excel conflict guard installation failed: {failure}"
        ) from failure

    source_unchanged = _sha256(source) == source_before
    if not source_unchanged:
        _safe_unlink(output)
        raise OutpatientConflictGuardPreviewError("Source workbook changed during preview build")

    return OutpatientConflictGuardPreviewReport(
        source_path=str(source),
        output_path=str(output),
        source_unchanged=True,
        outpatient_form_present=outpatient_form_present,
        popup_form_present=popup_form_present,
        popup_module_present=popup_module_present,
    )
