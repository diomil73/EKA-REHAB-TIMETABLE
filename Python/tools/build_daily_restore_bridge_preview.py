from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_change_ledger_vba import (  # noqa: E402
    LEDGER_SHEET,
    MODULE_NAME,
    VBA_CODE,
)


class DailyRestoreBridgePreviewError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def build_daily_restore_bridge_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> tuple[Path, int]:
    """Clone an authoritative daily workbook and refresh only its ledger VBA module.

    The source must already contain DAILY_CHANGE_LEDGER with real durable
    ReplacementIds. This deliberately does not reconstruct or rewrite ledger facts;
    it preserves the trusted workbook state and installs the current bridge code into
    a separate smoke-test copy.
    """

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise DailyRestoreBridgePreviewError(f"Source workbook not found: {source}")
    if source == output:
        raise DailyRestoreBridgePreviewError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise DailyRestoreBridgePreviewError("Source and output must both be .xlsm files")
    if not _has_vba(source):
        raise DailyRestoreBridgePreviewError("Source workbook has no VBA project")
    if output.exists() and not overwrite:
        raise DailyRestoreBridgePreviewError(f"Output already exists: {output}")
    if sys.platform != "win32":
        raise DailyRestoreBridgePreviewError("Windows with Microsoft Excel is required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyRestoreBridgePreviewError("pywin32 is required") from exc

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        try:
            output.unlink()
        except PermissionError as exc:
            raise DailyRestoreBridgePreviewError(
                f"Output workbook is open or locked by Excel: {output}"
            ) from exc

    shutil.copy2(source, output)
    excel = None
    workbook = None
    active_count = 0
    failure: Exception | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        try:
            ledger = workbook.Worksheets(LEDGER_SHEET)
        except Exception as exc:
            raise DailyRestoreBridgePreviewError(
                f"Source copy has no {LEDGER_SHEET} sheet"
            ) from exc

        last_row = int(ledger.Cells(ledger.Rows.Count, 1).End(-4162).Row)  # xlUp
        if last_row < 2:
            raise DailyRestoreBridgePreviewError("Daily change ledger has no fact rows")

        for row_index in range(2, last_row + 1):
            change_id = str(ledger.Cells(row_index, 1).Value or "").strip()
            status = str(ledger.Cells(row_index, 12).Value or "").strip().upper()
            if not change_id:
                continue
            if status == "ACTIVE":
                active_count += 1

        if active_count == 0:
            raise DailyRestoreBridgePreviewError(
                "Daily change ledger has no ACTIVE replacement rows for bridge smoke testing"
            )

        vbproject = workbook.VBProject
        _remove_component_if_present(vbproject, MODULE_NAME)
        module = vbproject.VBComponents.Add(1)
        module.Name = MODULE_NAME
        module.CodeModule.AddFromString(VBA_CODE)

        workbook.Save()
        excel.Run(f"'{workbook.Name}'!RefreshDailyChangeLedgerView")
        workbook.Save()
    except Exception as exc:
        failure = exc
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

    if failure is not None:
        output.unlink(missing_ok=True)
        if isinstance(failure, DailyRestoreBridgePreviewError):
            raise failure
        raise DailyRestoreBridgePreviewError(str(failure)) from failure

    if _sha256(source) != source_before:
        output.unlink(missing_ok=True)
        raise DailyRestoreBridgePreviewError("Source workbook changed during preview build")
    if not _has_vba(output):
        output.unlink(missing_ok=True)
        raise DailyRestoreBridgePreviewError("Preview lost its VBA project")

    return output, active_count


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Create a separate native Excel smoke-test copy of an authoritative daily "
            "workbook, preserving its ledger facts while refreshing modDailyChangeLedger "
            "to the current authoritative restore bridge implementation."
        )
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        output, active_count = build_daily_restore_bridge_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except DailyRestoreBridgePreviewError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DAILY RESTORE BRIDGE PREVIEW OK")
    print(f"Source: {args.source.resolve()}")
    print(f"Preview: {output}")
    print(f"Active ledger rows: {active_count}")
    print("Source unchanged: True")
    print("VBA preserved: True")
    print("Ledger bridge module refreshed: True")
    print("NEXT: open only this preview and exercise the intended Excel restore control.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
