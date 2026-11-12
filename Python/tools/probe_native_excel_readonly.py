"""Read-only native Excel COM probe for an isolated Diomil XLSM.

This probe is not a test-data importer and never asks Excel to Save.
No PatientIDs, sheet values, patient names, therapist names or workbook paths
are emitted to stdout. It disables VBA macros before Workbooks.Open and opens
with ReadOnly=True/UpdateLinks=0. Run on the local Windows Excel installation,
not GitHub-hosted CI, while the source is CLOSED in other Excel sessions.

Safety contract: only an isolated, institution-approved trial copy of the
workbook should be passed here. Not a production shared/live MASTER.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import sys
from zipfile import BadZipFile, ZipFile

REQUIRED = ("MASTER_SCHEDULE", "DAILY_INPUT", "PATIENTS", "THERAPIST_DAILY")
FORCE_DISABLE_MACROS = 3


class NativeExcelProbeError(RuntimeError):
    """Missing native contract, unexpected Excel state or file mutation."""


def digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as src:
        for piece in iter(lambda: src.read(1024 * 1024), b""):
            h.update(piece)
    return h.hexdigest()


def inspect_native_excel_readonly(
    path: Path, *, excel_factory=None,
) -> dict[str, object]:
    """Prove VBA-preserving native Excel read-back without ever saving.

    The optional factory exists solely for fully synthetic COM-contract tests.
    """
    source = Path(path).resolve()
    if source.suffix.lower() != ".xlsm" or not source.is_file():
        raise NativeExcelProbeError("Existing isolated .xlsm is required")
    if excel_factory is None:
        if sys.platform != "win32":
            raise NativeExcelProbeError("Native probe requires Windows Excel")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise NativeExcelProbeError("pywin32 is required on Windows") from exc
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    before = digest(source)
    try:
        with ZipFile(source) as zf:
            if "xl/vbaProject.bin" not in zf.namelist():
                raise NativeExcelProbeError("VBA project is missing")
        app = None
        book = None
        try:
            app = excel_factory()
            # Macros MUST be disabled before the file is opened. If the COM
            # property cannot be set, never open the workbook.
            app.AutomationSecurity = FORCE_DISABLE_MACROS
            app.DisplayAlerts = False
            app.EnableEvents = False
            app.ScreenUpdating = False
            app.AskToUpdateLinks = False
            book = app.Workbooks.Open(
                str(source), UpdateLinks=0, ReadOnly=True,
                AddToMru=False, IgnoreReadOnlyRecommended=True,
            )
            if not bool(book.ReadOnly):
                raise NativeExcelProbeError("Excel did not open workbook read-only")
            observed = {str(sheet.Name) for sheet in book.Worksheets}
            missing = tuple(name for name in REQUIRED if name not in observed)
            if missing:
                raise NativeExcelProbeError("Required operational sheets missing")
            # No cell reads: this inspection verifies only the workbook
            # and COM opening contract, not clinical record correctness.
            result = {
                "required_sheets_present": len(REQUIRED),
                "vba_present": True,
                "native_excel_readonly": True,
                "clinical_data_validated": False,
            }
        finally:
            if book is not None:
                try:
                    book.Close(SaveChanges=False)
                except Exception as exc:
                    raise NativeExcelProbeError("Excel read-only close failed") from exc
            if app is not None:
                try:
                    app.Quit()
                except Exception as exc:
                    raise NativeExcelProbeError("Excel process did not close") from exc
        return result
    except BadZipFile as exc:
        raise NativeExcelProbeError("Invalid .xlsm container") from exc
    finally:
        if not source.is_file() or digest(source) != before:
            raise NativeExcelProbeError("Source changed during read-only native probe")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Read-only native Excel COM proof for an isolated trial copy"
    )
    parser.add_argument("--isolated-copy", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = inspect_native_excel_readonly(args.isolated_copy)
    except Exception as exc:
        # Native COM exceptions may include sensitive workbook paths. Never
        # echo those exceptions, even when the probe was unsuccessful.
        print(f"SAFETY STOP: Native Excel read-only probe failed ({type(exc).__name__}).")
        return 2
    print("NATIVE EXCEL READ-ONLY PROBE OK")
    print(f"Expected operational sheets: {result['required_sheets_present']}")
    print("VBA present: YES")
    print("Opened read-only: YES")
    print("Input unchanged: YES")
    print("No clinical content checked; no patient programme imported or modified.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
