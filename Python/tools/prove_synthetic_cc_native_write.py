"""One-shot native Excel save/reopen proof on a fresh SYNTHETIC H10 CC copy.

Creates a new XLSM only. Neither the original synthetic CC nor prior probe
copies are modified. The only intended semantic addition is a separate
DIOMIL_NATIVE_PROBE worksheet; operational patient/history cells are not edited.

Never run this against an actual patient workbook. The native Excel probe is
run by an operator on Windows with Excel closed and macros forcibly disabled.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))
from zipfile import ZipFile

from tools.prepare_native_probe_copy import create_probe_copy, _sha256

PROBE_SHEET = "DIOMIL_NATIVE_PROBE"
PROBE_MARKER = "DIOMIL SYNTHETIC NATIVE COM SAVE REOPEN V1"


class NativeWriteProofError(RuntimeError):
    pass


def _require_synthetic_cc(source: Path, output: Path) -> None:
    if sys.platform != "win32":
        raise NativeWriteProofError("Native Excel write proof requires Windows")
    if not source.is_file() or not (
        "_DEMO_" in source.name.upper() and source.name.upper().endswith("_CC.XLSM")
    ):
        raise NativeWriteProofError("Only isolated synthetic H10 CC demo is permitted")
    if output.exists() or output.resolve() == source.resolve():
        raise NativeWriteProofError("Fresh distinct output is mandatory")
    if "_DEMO_" not in output.name.upper() or output.suffix.lower() != ".xlsm":
        raise NativeWriteProofError("Output must be a fresh synthetic DEMO XLSM")


def _session(excel_client, path: Path, *, readonly: bool):
    excel = excel_client.DispatchEx("Excel.Application")
    workbook = None
    try:
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(
            str(path), ReadOnly=readonly, UpdateLinks=0,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly) != readonly:
            raise NativeWriteProofError("Excel access mode differs from proof contract")
        return excel, workbook
    except Exception:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        raise


def native_write_reopen_proof(source: Path, output: Path, *, excel_client=None) -> dict:
    source, output = Path(source).resolve(), Path(output).resolve()
    _require_synthetic_cc(source, output)
    if excel_client is None:
        import win32com.client as excel_client  # type: ignore[import-not-found]
    source_hash = _sha256(source)
    copy_created = False
    phase = "copy"
    try:
        create_probe_copy(source, output)
        copy_created = True
        excel = book = None
        try:
            phase = "native-save"
            excel, book = _session(excel_client, output, readonly=False)
            if any(str(sheet.Name) == PROBE_SHEET for sheet in book.Worksheets):
                raise NativeWriteProofError("Diagnostic sheet already exists")
            sheet = book.Worksheets.Add()
            sheet.Name = PROBE_SHEET
            sheet.Range("A1").Value = PROBE_MARKER
            book.Save()
        finally:
            try:
                if book is not None:
                    book.Close(SaveChanges=False)
            finally:
                if excel is not None:
                    excel.Quit()
        excel = book = None
        try:
            phase = "native-reopen"
            excel, book = _session(excel_client, output, readonly=True)
            observed = book.Worksheets(PROBE_SHEET).Range("A1").Value
            if observed != PROBE_MARKER:
                raise NativeWriteProofError("Diagnostic marker did not persist")
        finally:
            if book is not None:
                book.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
        with ZipFile(output) as zf:
            if "xl/vbaProject.bin" not in zf.namelist():
                raise NativeWriteProofError("VBA project did not survive native save")
        if _sha256(source) != source_hash:
            raise NativeWriteProofError("Original CC changed during native proof")
        return {"copy_created": True, "marker_persisted": True,
                "vba_preserved": True, "source_unchanged": True}
    except Exception as exc:
        # Do not delete a failed native output: freeze it for diagnosis.
        # The original CC remains immutable by design.
        if not copy_created and output.exists():
            # Exclusive creation failed and an external file may own this path.
            pass
        raise NativeWriteProofError(f"Native proof stopped during {phase}") from exc
    finally:
        if source.is_file() and _sha256(source) != source_hash:
            raise NativeWriteProofError("Original synthetic CC unexpectedly changed")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="One-shot native save/reopen on synthetic CC copy")
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = native_write_reopen_proof(args.source, args.output)
    except Exception as exc:
        # Never print patient data, paths, COM internals or free-form Excel error text.
        print(f"SAFETY STOP: isolated native write proof failed ({type(exc).__name__}).")
        return 2
    if all(result.values()):
        print("SYNTHETIC NATIVE EXCEL SAVE REOPEN GREEN")
        print("New diagnostic sheet read-back: YES")
        print("VBA preserved: YES")
        print("Original synthetic CC unchanged: YES")
        print("This is not a patient scheduling change or import authorization.")
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
