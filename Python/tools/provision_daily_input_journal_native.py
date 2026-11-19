"""Provision DAILY_INPUT_JOURNAL in a disposable XLSM copy using native Excel.

The source workbook is never opened or modified. The output must be a new path.
The copy is saved by Excel, closed, reopened read-only, and independently checked
for VBA preservation, original worksheet visibility, and journal schema/protection.
No journal events are written and no operational hooks are installed.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import sys
from zipfile import BadZipFile, ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
for import_root in (REPO_ROOT, PYTHON_ROOT):
    if str(import_root) not in sys.path:
        sys.path.insert(0, str(import_root))

from Python.tools.prepare_native_probe_copy import create_probe_copy
from rehab_excel.daily_input_journal import JOURNAL_HEADERS, JOURNAL_SHEET
from rehab_excel.daily_input_journal_sheet import (
    ensure_daily_input_journal_sheet,
    read_native_journal_rows,
)


class NativeJournalProvisionError(RuntimeError):
    pass


FORCE_DISABLE_MACROS = 3
XL_SHEET_VERY_HIDDEN = 2


def _file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _vba_sha256(path: Path) -> str:
    try:
        with ZipFile(path) as archive:
            data = archive.read("xl/vbaProject.bin")
    except (BadZipFile, KeyError) as exc:
        raise NativeJournalProvisionError("XLSM VBA project is missing or invalid") from exc
    return sha256(data).hexdigest()


def _worksheet_snapshot(workbook) -> tuple[tuple[str, int], ...]:
    result = []
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets(index)
        result.append((str(sheet.Name), int(sheet.Visible)))
    return tuple(result)


def _verify_native_journal(workbook, expected_original_sheets: tuple[tuple[str, int], ...], expected_rows: tuple[tuple, ...]) -> None:
    observed = _worksheet_snapshot(workbook)
    original_names = {name for name, _ in expected_original_sheets}
    observed_original = tuple(item for item in observed if item[0] in original_names)
    if observed_original != expected_original_sheets:
        raise NativeJournalProvisionError("An original worksheet name/order/visibility changed")
    journal_matches = [item for item in observed if item[0] == JOURNAL_SHEET]
    if len(journal_matches) != 1 or journal_matches[0][1] != XL_SHEET_VERY_HIDDEN:
        raise NativeJournalProvisionError("Journal sheet is missing, duplicated, or not veryHidden")

    ws = workbook.Worksheets(JOURNAL_SHEET)
    if not bool(ws.ProtectContents):
        raise NativeJournalProvisionError("Journal worksheet is not protected")
    if tuple(ws.Cells(1, col).Value2 for col in range(1, 6)) != JOURNAL_HEADERS:
        raise NativeJournalProvisionError("Journal headers failed native reopen readback")
    if read_native_journal_rows(ws) != expected_rows:
        raise NativeJournalProvisionError("Journal history changed during native save/reopen")


def provision_daily_input_journal(
    source_path: Path,
    output_path: Path,
    *,
    excel_factory=None,
) -> dict[str, object]:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if source == output or source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise NativeJournalProvisionError("Distinct source and output .xlsm paths are required")
    if not source.is_file():
        raise NativeJournalProvisionError("Source workbook does not exist")
    if output.exists():
        raise NativeJournalProvisionError("Output already exists; refusing overwrite")
    if excel_factory is None:
        if sys.platform != "win32":
            raise NativeJournalProvisionError("Native journal provisioning requires Windows + Excel")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise NativeJournalProvisionError("pywin32 is required") from exc
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    source_hash = _file_sha256(source)
    source_vba_hash = _vba_sha256(source)
    created_copy = False
    app = None
    workbook = None
    try:
        create_probe_copy(source, output)
        created_copy = True
        if _file_sha256(source) != source_hash:
            raise NativeJournalProvisionError("Source changed during isolated-copy creation")

        app = excel_factory()
        # Disable macros before either writable or verification open.
        app.AutomationSecurity = FORCE_DISABLE_MACROS
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False

        workbook = app.Workbooks.Open(
            str(output), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly):
            raise NativeJournalProvisionError("Excel opened disposable copy read-only")
        baseline_sheets = _worksheet_snapshot(workbook)
        if any(name == JOURNAL_SHEET for name, _ in baseline_sheets):
            # Existing journal is validated by ensure(); do not replace or clear it.
            pass
        journal_sheet = ensure_daily_input_journal_sheet(workbook)
        expected_journal_rows = read_native_journal_rows(journal_sheet)
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None

        if _vba_sha256(output) != source_vba_hash:
            raise NativeJournalProvisionError("VBA project binary changed during journal provisioning")

        # Independent native reopen, read-only. Reopen does not execute macros.
        workbook = app.Workbooks.Open(
            str(output), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(workbook.ReadOnly):
            raise NativeJournalProvisionError("Verification reopen was not read-only")
        _verify_native_journal(workbook, baseline_sheets, expected_journal_rows)
        workbook.Close(SaveChanges=False)
        workbook = None
        app.Quit()
        app = None

        if _file_sha256(source) != source_hash:
            raise NativeJournalProvisionError("Source workbook changed during native provisioning")
        if _vba_sha256(output) != source_vba_hash:
            raise NativeJournalProvisionError("VBA project failed final preservation check")
        return {
            "copy_created": True,
            "source_unchanged": True,
            "vba_preserved": True,
            "native_save_reopen_verified": True,
            "journal_veryhidden_protected": True,
            "journal_empty": len(expected_journal_rows) == 0,
            "operational_hooks_installed": False,
        }
    except Exception:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if created_copy:
            output.unlink(missing_ok=True)
        raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Provision and verify an empty DAILY_INPUT_JOURNAL in a disposable XLSM copy"
    )
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = provision_daily_input_journal(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: Native journal provisioning failed ({type(exc).__name__}).")
        return 2
    print("DAILY_INPUT JOURNAL NATIVE PROVISION OK")
    print("Disposable copy created: YES")
    print("Source unchanged: YES")
    print("VBA project preserved: YES")
    print("Excel save/close/reopen readback: GREEN")
    print("Journal veryHidden + protected: YES")
    print("Journal empty: " + ("YES" if result["journal_empty"] else "NO (existing history preserved)"))
    print("Apply/Restore hooks installed: NO")
    print("This is a technical probe copy, not an operational workbook.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
