"""Privacy-safe, read-only inspection of installed native Apply/Restore VBA.

No macros, events, links, saves, log contents, patient values or paths emitted.
Requires Windows Microsoft Excel for the actual COM execution. The injectable
factory is used ONLY by fictional unit tests.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import sys
import tempfile

FORCE_DISABLE_MACROS = 3


class NativeContractInspectionError(RuntimeError):
    pass


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _module_code(vbproject, name: str) -> str:
    module = vbproject.VBComponents(name).CodeModule
    count = int(module.CountOfLines)
    if count < 0:
        raise NativeContractInspectionError("Invalid native VBA module")
    return str(module.Lines(1, count)) if count else ""


def _log_status(name: str) -> str:
    """Never open patient-bearing diagnostic logs or echo their contents."""
    path = Path(tempfile.gettempdir()) / name
    return "PRESENT" if path.is_file() else "MISSING"


def inspect_workbook(path: Path, *, excel_factory=None) -> dict[str, bool]:
    source = Path(path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeContractInspectionError("Existing isolated XLSM required")
    if excel_factory is None:
        if sys.platform != "win32":
            raise NativeContractInspectionError("Windows + Microsoft Excel required")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise NativeContractInspectionError("pywin32 required") from exc
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    initial_digest = _digest(source)
    excel = None
    workbook = None
    try:
        excel = excel_factory()
        # Set AutomationSecurity BEFORE Workbooks.Open. Never run workbook VBA.
        excel.AutomationSecurity = FORCE_DISABLE_MACROS
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(
            str(source), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(workbook.ReadOnly):
            raise NativeContractInspectionError("Native inspection must be read-only")
        vbproject = workbook.VBProject
        daily = _module_code(vbproject, "modDailyInputAction")
        ledger = _module_code(vbproject, "modDailyChangeLedger")

        checks = {
            "daily.character_apply": "Private Sub ApplyCharacterStrike" in daily,
            "daily.character_verify": "Private Function VerifyCharacterStrike" in daily,
            "daily.status_normalizer": "NormalizeMasterPatientStatusPresentation master" in daily,
            "daily.status_character_clear": "ApplyCharacterStrike master.Cells(row, statusCol), False" in daily,
            "ledger.canonical_apply": "ApplyValidatedDailyInputPayload(patchPath)" in ledger,
            "ledger.old_apply_absent": "ApplyPayloadFileInPlace" not in ledger,
            "ledger.safe_anchor": 'wsDaily.Cells.Find(What:="*"' in ledger,
            "ledger.old_broad_clear_absent": 'startRow + 200).ClearContents' not in ledger,
            "ledger.change_id_declared": "Dim changeId As String" in ledger,
        }
        return checks
    except NativeContractInspectionError:
        raise
    except Exception:
        # COM errors can embed workbook paths or clinical data. Never forward.
        raise NativeContractInspectionError("Native VBA read-only contract inspection failed") from None
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                if not source.is_file() or _digest(source) != initial_digest:
                    raise NativeContractInspectionError(
                        "Source changed during read-only VBA inspection"
                    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Inspect isolated native Apply/Restore VBA safely, no workbook writes"
    )
    parser.add_argument("--input", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        checks = inspect_workbook(args.input)
    except Exception as exc:
        print(
            "SAFETY STOP: Native Apply/Restore read-only contract failed "
            f"({type(exc).__name__})."
        )
        return 2

    print("NATIVE APPLY/RESTORE CONTRACT INSPECTION")
    for key, passed in checks.items():
        print(f"{key}={'PASS' if passed else 'FAIL'}")
    print("Local DAILY INPUT log: " + _log_status("eka_daily_input_last_error.log"))
    print("Local RESTORE log: " + _log_status("eka_daily_restore_last_error.log"))
    print("Source unchanged: YES; no macros executed, no workbook saved.")
    print("Diagnostic/static VBA contract only; not Apply/Restore functional acceptance.")
    return 0 if all(checks.values()) else 3


if __name__ == "__main__":
    raise SystemExit(main())
